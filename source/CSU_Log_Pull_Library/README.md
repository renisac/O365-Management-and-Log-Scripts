# Office 365 Log Pull Library

## Use case

Office 365 Log Pull Library is a library written in C# to help applications interact with the Office 365 Management API to manage logs. The project was started as a way to re-establish logging user access for Exchange after a migration from an on-premises Exchange Server to Exchange Online.

This library helps applications

* Authenticate to the Microsoft Management Activty API 
* Start, Stop, and List subscriptions
* Retrieve available activity log content

When used with [ExchangeOnlineLogPull](../CSU_EO_Log_Pull) Console Application it provides a quick and easy way to retrieve activity logs from online services such as Exchange, SharePoint, and Azure Active Directory.

## Brief description

In a command prompt, navigate to "TestApplication\bin\Debug" or "TestApplication\bin\Release" (Depending on your settings when building the solution), then run "TestApplication.exe".

TestApplication.exe tries to list all subscriptions for the tenant. It will most likely return an empty list, but this allows you to test if authentication is working properly.

Check out [ExchangeOnlineLogPull](../CSU_EO_Log_Pull) Console Application for more features or to see examples of code which interacts with the library.

## Data sensitivity

None

## Prerequisites

* Office 365 tenant admin account
* X.509 certificate

### Installing the library

Open the "ISOLogPullLibrary.sln" in Visual Studio, right-click "Solution" in Solution Explorer, select "Build Solution."

### Testing the library

To use the library you will need to:

* Register the Application with Azure AD
* Specify permissions for Application
* Install certificate on local machine
* Register certificate with Azure AD
* Add necessary information to the AppOptions.config file 
* Run a console application calling the library

#### Registering the Application and Specifying permissions

These two steps require Microsoft tenant admin credentials, but are fairly straight forward. 
Follow Microsoft's [Get Started with Office 365 Management APIs](https://msdn.microsoft.com/en-us/office-365/get-started-with-office-365-management-apis) guide

To obtain logs for:

* Azure Active Directory
* Exchange
* SharePoint
* General

We specified the following permissions:

* Microsoft Graph
  * Read all users' full profiles
    * Read directory data
    * Read all group
* Office 365 SharePoint Online

#### Install certificate on local machine

A self-signed certificate can be used to authenticate to the Microsoft APIs. The guide linked to above provides instrcution for creating and installing the certificate using Windows SDK makecert and Powershell.
Below I have listed instructions for using openssl to create a certificate and Windows mmc to install it.

Open a bash prompt and run the following commands:

```cmd
openssl genrsa -out <preferred_name>.key 2048
```

This will generate a key pair to be used in the certificate:

```cmd
openssl req -new -key <preferred_name>.key -out <preferred_name>.csr
```

Fill out the relevant information. Leave the password blank. Now we will be able to sign the certificate

```cmd
openssl x509 -req -days 366 -in <preferred_name>.csr -signkey <preferred_name> -out <preferred_name>.crt
```

Now we just need to generate the .pfx file:

```cmd
openssl pkcs12 -export -nodes -out <preferred_name>.pfx -inkey <preferred_name>.key -in <preferred_name>.crt 
```

You should now have a .pfx file which can be installed in your computers cert store.

To install the cert on the local machine:

* On the Start Menu click Run an type "mmc" then hit enter
* Click the file tab, then select "Add/Remove Snap-in"
* In the left-column, under "Available snap-ins:", find and select "Certifictes", click "Add >", click "OK"
* In the far left column, expand "Certificates", then expand "Personal", right-click "Certificates", select "All Tasks"->"Import"
  * Once in Certificate Import Wizard, select "Local Machine", click "Next", then click "Browse" and select your <preferred_name>.pfx file

#### Register certificate with Azure AD

For this refer to Microsoft's [Get Started with Office 365 Management APIs](https://msdn.microsoft.com/en-us/office-365/get-started-with-office-365-management-apis) guide again.
Jump to the "Configure an X.509 certificate to enable service-to-service calls" section, step 6 and 7. 

#### Add necessary information to the AppOptions.config file

The AppOptions.config file is a config file that stores information like Tenant ID, Client ID, and Certificate Thumbprint.

##### Important - this file is located in the directory of the executable which calls the library, not the directory of the library itself

The easiest way to fill out this file is to run the executable which will create the file with default settings stored in file. (Bug - the executable may hang, just kill the process)
This repository has a TestApplication executable which can be run to test the library. In a command prompt, navigate to "TestApplication\bin\Debug" or "TestApplication\bin\Release" (Depending on your settings when building the solution), then run "TestApplication.exe".

Open the AppOptions.config file in a text editor and add the values for any blank fields. The file should appear similar to below:

```txt
aadinstance=https://login.microsoftonline.com  
tenant=  
clientid=  
tenantid=  
certthumbprint=  
resourceid=https://manage.office.com  
subscriptiontype=exchange  
tempfolder=C:\Users\<current_user>\AppData\Local\Temp\
```

### License and/or permissions required

A1 or higher

### Known issues

None

### Acknowledgements and Credits

California State University, Sacramento