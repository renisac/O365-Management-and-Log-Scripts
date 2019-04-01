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
 * Notes: Only works with certificate authentication in this version
 *         
 * *****************************/


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;


using Microsoft.IdentityModel.Clients.ActiveDirectory;

// Version 2 - add ability to authenticate with clientSecret or username and pass
//             would need to store both of these in an encrypted file or it would require user interaction

namespace ISOLogPullLibrary
{

    // This class will attempt to authenticate using a certificate on instantiation
    // and then store the results
    public class Authentication
    {
        // Authentication Result will have the access token we need to make requests to the api 
        public AuthenticationResult Result
        {
            get; private set;
        }

        #region CertAuth

        // Need async method to use await for authentication
        // Created method CertAuthentication to handle async call

        // Constructor
        // If authentication fails it return null
        public Authentication()
        {
            try
            {
                Result = CertAuthentication().Result;
            }
            catch
            {
                Result = null;
            }
            
            
        }

        // Method which attempts to authenticate with cert
        private async Task<AuthenticationResult> CertAuthentication()
        {
            // Create an instance of AppOptions to get values from config
            ISOLogPullLibrary.AppOptions appOptions = new AppOptions();
            ClientAssertionCertificate cac;
            AuthenticationResult result = null;
            AuthenticationContext authContext = null;

            try
            {
                //this happens regardless of authentication method
                // Find Certificate by thumbprint in Cert Store using getCert method
                // getCert handles checking different stores for cert and permissions
                X509Certificate2 cert = getCert(appOptions.CertThumbPrint);

                // Use clientId and Certificate to create Client Assertion
                cac = new ClientAssertionCertificate(appOptions.ClientId, cert);

                // determine if we're using oauth and act accordingly
                string oAuthUri = appOptions.Authority + appOptions.Oauth2AuthorityId;
                if (appOptions.UseOauth2)
                {
                    authContext = new AuthenticationContext(oAuthUri);
                }
                else
                {
                    // Create an Authentication Context
                    // Uses the Authority attribute from AppOptions
                    // This is the AADInstance + Tenant and is not actually listed in config file
                    authContext = new AuthenticationContext(appOptions.Authority);
                }
                
                
                //Attempt to Authenticate to resource id
                try
                {
                    if (appOptions.UseOauth2)
                    {
                        IPlatformParameters ipm = new PlatformParameters(PromptBehavior.Never);
                        //result = await authContext.AcquireTokenAsync(appOptions.ResourceId, appOptions.ClientId, new Uri(appOptions.Oauth2RedirectUi), ipm,);
                        
                        result = await authContext.AcquireTokenAsync(appOptions.ResourceId, cac);
                    }
                    else
                    {
                        // Use the acquire token method and wait until finishes
                        result = await authContext.AcquireTokenAsync(appOptions.ResourceId, cac);
                        
                    }
                    return result;
                }
                catch(Exception ex)
                {
                    // Return null if can't authenticate
                    Debug.WriteLine(ex.Message);
                    return null;
                }
                            

            }
            catch (AdalException ex)
            {

                if (ex.ErrorCode != "user_interaction_required")
                {
                    // An unexpected error occurred.
                    Debug.WriteLine(ex.Message);
                }

                // An unexpected error occurred, or user canceled the sign in.
                if (ex.ErrorCode != "access_denied"){
                    Debug.WriteLine(ex.Message);
                }

                return null;
            }
            #endregion
        }


        // Get Cert method will search cert stores for the certificate
        // Will first check local machine then current user cert stores
        // If certificate installed on Local Machine application will need to run as Admin
        // In either case, Certificate should be placed in Personal folder
        private X509Certificate2 getCert(string certThumbPrint)
        {
            X509Store certStore;
            X509Certificate2Collection certCollection;
            X509Certificate2 cert;

            // Get cert store from local machine and open
            certStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            certStore.Open(OpenFlags.ReadOnly);

            //Returns collection of Certs based on thumbprint we grab the first one and store in cert
            certCollection = certStore.Certificates.Find(X509FindType.FindByThumbprint, certThumbPrint, false);
            cert = null;

            // If any certs were returned grab first one
            // If it does not enter if statement cert will remian null
            if (certCollection.Count > 0)
            {
                cert = certCollection[0];
            }
            
            // Even if user does not have permissions for the certificate private key 
            // but it is installed on Local Machine the above commands will retrieve cert
            // Here we test if user has proper permissions
            // If they do return cert otherwise check in Current User store
            // This will also handle if no Certs matching thumprint were found in the Local Machine store   
            try
            {
                var test = cert.PrivateKey;
                return cert;
            }
            catch
            {
                // Reach here if cert not in Local Machine or user does not have proper permissions

                // Get Current User store and open
                certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                certStore.Open(OpenFlags.ReadOnly);

                //Returns collection of Certs based on thumbprint we grab the first one and store in cert
                certCollection = certStore.Certificates.Find(X509FindType.FindByThumbprint, certThumbPrint, false);
                cert = null;

                // If any certs were returned grab first one
                // If it does not enter if statement cert will remian null
                if (certCollection.Count > 0)
                {
                    cert = certCollection[0];
                }

                // Test if user has permissions or if no certs were found
                try
                {
                    var test = cert.PrivateKey;
                    return cert;
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
