# MAPIToolkit

Allow managing MAPI profile and running MAPI operations against mailobxes. 

## Help information

**Usage:**
```sh
 [-?]
 [-action               {addservice, listservice, listallservices, removeservice, removeallservices, updateservice}]
 [-configfilepath       <string>]
 [-customsearchbase     <string>]
 [-defaultsearchbase    {true, false}]
 [-displayname          <string>]
 [-enablebrowsing       {true, false}]
 [-logfilepath          <string>]
 [-loggingmode          {none, console, file, all, debug}]
 [-maxentries           <int>]
 [-newdisplayname       <string>]
 [-newservername        <string>]
 [-newserverport        <int>]
 [-password             <string>]
 [-preservecase]
 [-profilemode          {default, specific, all}]
 [-profilename          <string>]
 [-registry]
 [-requirespa           {true, false}]
 [-saveconfig           {true, false}]
 [-searchtimeout        <int>]
 [-servername           <string>]
 [-serverport           <int>]
 [-servicetype          {addressbook}]
 [-setsearchindex       {<int>}]
 [-username             <string>]
 [-usessl               {true, false}]
```
**Options:**
```sh
 -?                   : Displays the help info.
 -action              : Action(s) to perform.
 -configfilepath      : Path to the input configuration file.
 -customsearchbase    : custom search base in case defaultsearchbase is set to false.
 -defaultsearchbase   : If "true" the default search base is to be used. The default value is 'true'.
 -displayname         : The display name of the service to run the action(s) against.
 -enablebrowsing      : Indicates whether browsing the address book contens is supported.
 -logfilepath         : Path towards the log file where informatin is to be logged.
 -loggingmode         : Indicates how logging is captured.
 -maxentries          : The maximum number of results returned by a search in the address book. The default value is 100.
 -newdisplayname      : Display name to replace the current display name of the service with.
 -newservername       : Server name to replace the current server name with in the speciifed service.
 -newserverport       : Server port to replace the current server port with in the speciifed service.
 -password            : The password to use for authenticating. This must be a clear text passord. It will be encrypted via CryptoAPI and stored in the address book settings.
 -preservecase        : Indicates whether the string values passed for service or server names are to be processed as typed, preserving the case instead of transposing the values to lowercase.
 -profilemode         : Indicates whether to run the action on all profiles or a specific profile.
 -profilename         : Indicates the name of the profile to run the action against. If left empty, the default profile will be used, unles the profilemode specified is "all".
 -registry            : Indicates whether to read the configuration from the registry if previously saved with "-saveconfig true".
 -requirespa          : "true" if Secure Password Authentication is required is required. The default value is "false"
 -saveconfig          : Indicates whether to save the current configuration in teh registry or no
 -searchtimeout       : The number of seconds before the search request times out. The default value is 60 seconds.
 -servername          : The LDAP address book server address. For example "ldap.contoso.com".
 -serverport          : The LDAP port to connect to. The standard port for Active Directory is 389.
 -servicetype         : Indicates the type of service to run the action against.
 -setsearchindex      : Sets the address book search order in Outlook as such as the address book being added is used based on the index specified. For example, for a value of 1, the new address book will be the first one in the search list.
 -username            : The Username to use for authenticating in the form of domain\username, UPN or just the username if domain name not applicable or not required. Leave blank if a username and password are not required.
 -usessl              : "true" if a SSL connection is required.The default value is "false".
 ```
 
## Listing specific Address Book services

### Example 1: 
**Syntax** 
```sh
MAPIToolkitConsole.exe -action listservice -servicetype addressbook -servername ldap.contoso.com
```
**Sample Output**
```sh
2019.8.14 11:50:42 INFO Action listservice will run against 1 services
2019.8.14 11:50:42 INFO   Display Name        : Contoso
2019.8.14 11:50:42 INFO   Ldap Server Name    : ldap.contoso.com
2019.8.14 11:50:42 INFO   Ldap Server Port    : 389
2019.8.14 11:50:42 INFO   Username            :
2019.8.14 11:50:42 INFO   Search Base         :
2019.8.14 11:50:42 INFO   Search Timeout      : 60
2019.8.14 11:50:42 INFO   Maximum entries     : 100
2019.8.14 11:50:42 INFO   Use SSL             : false
2019.8.14 11:50:42 INFO   Require SPA         : false
2019.8.14 11:50:42 INFO   Enable browsing     : false
2019.8.14 11:50:42 INFO   Default search base : true
2019.8.14 11:50:42 SUCCESS Address book service succesfully listed
2019.8.14 11:50:42 SUCCESS Action succesfully run on profile: microsoft
```

### Example 2: 
**Syntax** 
```sh
MAPIToolkitConsole.exe -action listservice -servicetype addressbook -displayname Contoso
```
**Sample Output**
```sh
2019.8.14 11:51:7 INFO Action listservice will run against 1 services
2019.8.14 11:51:7 INFO   Display Name        : Contoso
2019.8.14 11:51:7 INFO   Ldap Server Name    : ldap.contoso.com
2019.8.14 11:51:7 INFO   Ldap Server Port    : 389
2019.8.14 11:51:7 INFO   Username            :
2019.8.14 11:51:7 INFO   Search Base         :
2019.8.14 11:51:7 INFO   Search Timeout      : 60
2019.8.14 11:51:7 INFO   Maximum entries     : 100
2019.8.14 11:51:7 INFO   Use SSL             : false
2019.8.14 11:51:7 INFO   Require SPA         : false
2019.8.14 11:51:7 INFO   Enable browsing     : false
2019.8.14 11:51:7 INFO   Default search base : true
2019.8.14 11:51:7 SUCCESS Address book service succesfully listed
2019.8.14 11:51:7 SUCCESS Action succesfully run on profile: microsoft
```

## Listing all Address Book services

### Example 1: 
**Syntax** 
```sh
MAPIToolkitConsole.exe -action listallservices -servicetype addressbook
```
**Sample Output**
```sh
2019.8.14 11:48:33 INFO   Listing entry #0
2019.8.14 11:48:33 INFO   Display Name        : ldap.contoso.com
2019.8.14 11:48:33 INFO   Ldap Server Name    : ldap.contoso.com
2019.8.14 11:48:34 INFO   Ldap Server Port    : 389
2019.8.14 11:48:34 INFO   Username            :
2019.8.14 11:48:34 INFO   Search Base         :
2019.8.14 11:48:34 INFO   Search Timeout      : 60
2019.8.14 11:48:34 INFO   Maximum entries     : 100
2019.8.14 11:48:34 INFO   Use SSL             : false
2019.8.14 11:48:34 INFO   Require SPA         : false
2019.8.14 11:48:34 INFO   Enable browsing     : false
2019.8.14 11:48:34 INFO   Default search base : true
2019.8.14 11:48:34 SUCCESS Address book services succesfully listed
2019.8.14 11:48:34 SUCCESS Action succesfully run on profile: microsoft
```

## Updating existing Address Book services

### Example 1: 
**Syntax**
```sh
MAPIToolkitConsole.exe -action updateservice -servername ldap.contoso.com -newdisplayname "Contoso" -servicetype addressbook
```
**Sample Output**
```sh
2019.8.14 11:49:54 INFO Action updateservice will run against 1 services
2019.8.14 11:49:54 SUCCESS Address book service succesfully updated
2019.8.14 11:49:54 SUCCESS Action succesfully run on profile: microsoft
```

### Example 2: 
**Syntax**
```sh
MAPIToolkitConsole.exe -action updateservice -displayname Fabrikam -newserverport 389 -servicetype addressbook
```
**Sample Output**
```sh
2019.8.14 11:55:14 INFO Action updateservice will run against 1 services
2019.8.14 11:55:14 SUCCESS Address book service succesfully updated
2019.8.14 11:55:14 SUCCESS Action succesfully run on profile: microsoft
```
### Example 3: 
**Syntax**
```sh
MAPIToolkitConsole.exe -action updateservice -newdisplayname ldap.contoso.com -registry -saveconfig true
```
**Sample Output**
```sh
2019.8.14 14:37:48 INFO Action updateservice will run against 1 services
2019.8.14 14:37:48 SUCCESS Address book service succesfully updated
2019.8.14 14:37:48 SUCCESS Action succesfully run on profile: microsoft
```

## Adding new Address Book services

### Example 1: 
**Syntax**
```sh
MapiToolkitConsole.exe -action addservice -servicetype addressbook -displayname Fabrikam -servername ldap.fabrikam.com -serverport 636
```
**Sample Output**
```sh
2019.8.14 11:52:29 SUCCESS Address book service succesfully added 
2019.8.14 11:52:29 SUCCESS Action succesfully run on profile: microsoft
```
### Example 2:
**Syntax**
```sh
MAPIToolkitConsole.exe -action addservice -servicetype addressbook -displayname tailspintoys.com -servername ldap.tailspintoys.com
```
**Sample Output**
```sh
2019.8.14 11:54:4 SUCCESS Address book service succesfully added 
2019.8.14 11:54:4 SUCCESS Action succesfully run on profile: microsoft
```
### Example 3: 
**Syntax**
```sh
MAPIToolkitConsole.exe -action addservice -servicetype addressbook -configfilepath C:\Temp\Configuration.xml
```
**Sample Output**
```sh
2019.8.14 13:40:19 SUCCESS Address book service succesfully added
2019.8.14 13:40:19 SUCCESS Action succesfully run on profile: microsoft
```
### Example 4:
**Syntax**
```sh
MAPIToolkitConsole.exe -action addservice -servicetype addressbook -configfilepath C:\Temp\Configuration.xml -saveconfig true
```
**Sample Output**
```sh
2019.8.14 13:42:59 SUCCESS Address book service succesfully added
2019.8.14 13:42:59 SUCCESS Action succesfully run on profile: microsoft
```
## Removing specific Address Book services                                                             

### Example 1: 
**Syntax**
```sh
MAPIToolkitConsole.exe -action removeservice -servername ldap.fabrikam.com -servicetype addressbook
```
**Sample Output**
```sh
2019.8.14 12:22:50 INFO Action removeservice will run against 1 services
2019.8.14 12:22:50 SUCCESS Address book service succesfully removed
2019.8.14 12:22:50 SUCCESS Action succesfully run on profile: microsoft
```
### Example 2: 
**Syntax**
```sh
MAPIToolkitConsole.exe -action removeservice -displayname Fabrikam -servicetype addressbook
```
**Sample Output**
```sh
2019.8.14 12:50:9 INFO Action removeservice will run against 1 services
2019.8.14 12:50:9 SUCCESS Address book service succesfully removed
2019.8.14 12:50:9 SUCCESS Action succesfully run on profile: microsoft
```

## Removing all Address Book services 

### Example 1:
**Syntax**
```sh
MAPIToolkitConsole.exe -action removeallservices -servicetype addressbook
```
**Sample Output**
```sh
2019.8.14 12:51:29 INFO Action removeallservices will run against 2 services
2019.8.14 12:51:29 INFO Number of services found: 2
2019.8.14 12:51:29 SUCCESS Action succesfully run on profile: microsoft
```

## Sample configuration XML file
```sh
<?xml version="1.0"?>
<ABConfiguration>
    <!--DisplayName-->
    <!--The name displayed in the AddressBook menu (form).-->
    <DisplayName>Contoso</DisplayName>
    <!--ServerName-->
    <!--The LDAP address book server address. For example "ldap.contoso.com"-->
    <ServerName>ldap.contoso.com</ServerName>
    <!--ServerPort-->
    <!--The LDAP port to connect to. The standard port for Active Directory is 389.-->
    <ServerPort>389</ServerPort>
    <!--UseSSL-->
    <!--"True" if a SSL connection is required. The default value is "False".-->
    <UseSSL>False</UseSSL>
    <!--UserName-->
    <!--The Username to use for authenticating in the form of domain\username, UPN or just the username if domain name not applicable or not required. Leave blank if a username and password are not required.-->
    <Username></Username>
    <!--Password-->
    <!--The Password to use for authenticating. This must be a clear text passord. It will be encrypted via CryptoAPI and stored in the AB settings.-->
    <Password></Password>
    <!--RequireSecurePasswordAuth-->
    <!--"True" if Secure Password Authentication is required is required. The default value is "False".-->
    <RequireSecurePasswordAuth>False</RequireSecurePasswordAuth>
    <!--SearchTimeoutSeconds-->
    <!--The number of seconds before the search request times out. The default value is 60 seconds.-->
    <SearchTimeoutSeconds>60</SearchTimeoutSeconds>
    <!--MaxEntriesReturned-->
    <!--The maximum number of results returned by a search in this AB. The default value is 100.-->
    <MaxEntriesReturned>100</MaxEntriesReturned>
    <!--DefaultSearchBase-->
    <!--"True" the default search base is to be used. The default value is "True".-->
    <DefaultSearchBase>True</DefaultSearchBase>
    <!--CustomSearchBase-->
    <!--Custom search base in case DefaultSearchBase is set to False.-->
    <CustomSearchBase></CustomSearchBase>
    <!--EnableBrowsing-->
    <!--Indicates whether browsing the AB contens is supported. -->
    <EnableBrowsing>False</EnableBrowsing>
</ABConfiguration>
```

## Information about this sample

MAPIToolkit uses Extended MAPI to access and manage the Outlook (MAPI) profile configuration. 

If running with the **-saveconfig true** switch, the tool writes the configuration settings under *HKEY_CURRENT_USER\SOFTWARE\Microsoft Ltd\MAPIToolkit*. Address book specific configuration is stored under *HKEY_CURRENT_USER\SOFTWARE\Microsoft Ltd\MAPIToolkit\AddressBook*.

For example:
```s
MAPIToolkitConsole.exe -action addservice -displayname ldap.contoso.com -saveconfig true
```

If you wish to reuse the previous configuraiton and only edit a single parameter or a small number of parameters, you can use the **-registry** switch which will read the previously saved configuration and specify any overrides to the saved configuration. If you wish for the overrides to be saved to the registry you will have to use **-saveconfig true** again. 

For example:
```s
MAPIToolkitConsole.exe -action addservice -newdisplayname Contoso -registry -saveconfig true
```

The sample also uses some default values to fill in any gaps in the service configuration. These defaults are as follows:
```s
loggingmode:       console
profilemode:       default
saveconfig:        false
serverport:        389
usessl:            false
requirespa:        false
searchtimeout:     60
maxentries:        100
defaultsearchbase: true
enablebrowsing:    false
```

## Disclaimer

This is a code sample maintained by myself alone. If you run into any problems please log the issues here on github. Please do not raise any support cases with Microsoft Customer Service and Support as Microsoft has no obligation to provide support for this sample, nor will it. 
