/////////////////////////////////////////////////////////////////
// Copyright (c) Microsoft Corporation.  All rights reserved
//
// Title:	Lib_CommonGPMCFunctions.js
// Author:	mtreit@microsoft.com
// Created:	7/16/2002
// Purpose:	Provides a library of common helper functions
//		for use when scripting the GPMC interfaces.
//
//		This library must be included with the sample
//		WSH scripts that ship with the GPMC
/////////////////////////////////////////////////////////////////

///////////////////////////////////////
// Initialization
///////////////////////////////////////

// Create global objects for use by all of the functions
var GPM = new ActiveXObject("GPMgmt.GPM");
var Constants = GPM.GetConstants();

///////////////////////////////////////
// Common Function Library
///////////////////////////////////////

//
// Note: The functions in this section are shared by
// many of the GPMC sample scripts. This section may be
// pasted directly in each individual script to ensure they
// will work stand-alone, or may be collected in a library
// file and accessed using the 'include' functionality  
// provided by the WSF script format.
// 

// Takes a GPO name or GUID and returns the GPO
function GetGPO(szGPOName, GPMDomain)
{
	var GPO;

	// Get the GPO object for the specified GPO
	try
	{
		GPO = GPMDomain.GetGPO(szGPOName);
	}
	catch (err)
	{
		// The attempt to get the GPO failed. The user may have
		// passed in the name instead of GUID, so fetch by name.
		try
		{
			GPO = GetGPOByName(szGPOName, GPMDomain);
		}
		catch (err)
		{
			WScript.Echo("Could not find GPO " + szGPOName);
			return false;
		}
	}

	return GPO;

}


// Given a GPO name or ID (GUID), returns that GPO from the directory.
// If no GPO is found, returns null
// If multiple GPOs exist by that name, returns the resulting collection
//
function GetGPOByName(szGPOName, GPMDomain)
{
	// Create a search criteria object for the name
	var GPMSearchCriteria = GPM.CreateSearchCriteria();
	GPMSearchCriteria.Add(Constants.SearchPropertyGPODisplayName, Constants.SearchOpEquals, szGPOName);

	// Search for the specified GPO
	var GPOList = GPMDomain.SearchGPOs(GPMSearchCriteria);

	if (GPOList.Count == 0)
	{
		return false; // No GPO found
	}

	// The following could return a collection of multiple GPOs if more than one GPO
	// with the same name exists in the domain
	//
	if (GPOList.Count == 1)
	{
		return GPOList.Item(1);
	}
	else
	{
		return GPOList;
	}

}

// Retrieves the WMI filter with the specified name
function GetWMIFilter(szWMIFilterName, GPMDomain)
{
	var GPMSearchCriteria = GPM.CreateSearchCriteria();
	var FilterList = GPMDomain.SearchWMIFilters();
	var e = new Enumerator(FilterList);
	var WMIFilter;

	for (; !e.atEnd(); e.moveNext())
	{
		WMIFilter = e.item();
		if (WMIFilter.Name.toLowerCase() == szWMIFilterName.toLowerCase())
		{
			return WMIFilter;
		}
	}

	return false;
}

// Attempts to retrieve a SOM by name or path from the directory. Will return a single GPMSOM object, or
// an array of such objects if more than one with the given name is found.
//
function GetSOM(szSOMName, GPMDomain)
{

	// Check if this is the domain level - if so, get the SOM for the domain and return it
	if (szSOMName.toLowerCase() == GPMDomain.Domain.toLowerCase())
	{
		return GPMDomain.GetSOM(""); // Returns the SOM representing the domain
	}

	// First try to get the SOM, in case a valid LDAP-style path was passed in
	try
	{
		var GPMSOM = GPMDomain.GetSOM(szSOMName);
	}
	catch (err)
	{
		try
		{
			// Might be a site instead of a domain or oU
			GPMSOM = GPMSitesContainer.GetSite(szSOMName);
		}
		catch (err)
		{
			GPMSOM = false;
		}
	}

	if (GPMSOM)
	{
		return GPMSOM;
	}
	
	// Search for the SOM by name, using ADSI

	// Create an array to hold the results, as we may find more than one SOM with the specified name
	var aResult = new Array();
	
	// Define ADS related values - see IADS.h
	var ADS_SCOPE_BASE = 0;
	var ADS_SCOPE_ONELEVEL = 1;
	var ADS_SCOPE_SUBTREE = 2;
	var ADSIPROP_CHASE_REFERRALS		=	0x9;
	var ADS_CHASE_REFERRALS_NEVER		=	0;
	var ADS_CHASE_REFERRALS_SUBORDINATE	=	0x20;
	var ADS_CHASE_REFERRALS_EXTERNAL	=	0x40;
	var ADS_CHASE_REFERRALS_ALWAYS		=	ADS_CHASE_REFERRALS_SUBORDINATE | ADS_CHASE_REFERRALS_EXTERNAL;

	var szLDAPSuffix = GPMDomain.GetSOM("").Path;

	// Create the ADO objects and open the connection
	var ADOConnection = new ActiveXObject("ADODB.Connection");
    	var ADOCommand =  new ActiveXObject("ADODB.Command");
	ADOConnection.Provider = "ADsDSOObject";    
	ADOConnection.Open("Active Directory Provider");    
	ADOCommand.ActiveConnection = ADOConnection;
	
	// First look for OUs
	var szDomainLDAPPath = "LDAP://" + szLDAPSuffix;
	var szSQL = "select AdsPath from '" + EscapeString(szDomainLDAPPath) + "'";
	szSQL += " where Name='" + szSOMName + "'";

	// Execute the search
	ADOCommand.CommandText = szSQL;
	ADOCommand.Properties("Page Size") = 1000;
	ADOCommand.Properties("Timeout") = 500;
	ADOCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE;
	ADOCommand.Properties("Cache Results") = false;
	ADOCommand.Properties("Chase Referrals") = ADS_CHASE_REFERRALS_EXTERNAL; // Needed when querying a different domain

	try
	{
		var rs = ADOCommand.Execute();
	}
	catch (err)
	{
		WScript.Echo("There was an error executing the DS query " + szSQL);
		WScript.Echo("The error was:");
		WScript.Echo(ErrCode(err.number) + " - " + err.description);
		return false;
	}

	var SOM;
	while ( ! rs.eof )
	{
		SOM = GetObject(rs.Fields(0));
		
		// Ignore objects that are not OUs or the domain level
		if (SOM.Class == 'organizationalUnit' || SOM.Class == 'fTDfs')
		{
			GPMSOM = GPMDomain.GetSOM(SOM.ADsPath)
			aResult = aResult.concat(GPMSOM);
		}
		
		rs.MoveNext();
	}

	// Get the LDAP suffix from the forest name
	ForestDomain = GPM.GetDomain(szForestName, "", Constants.UseAnyDC);
	szLDAPSuffix = ForestDomain.GetSOM("").Path;

	var szSitesLDAPPath = "LDAP://CN=Sites,CN=Configuration," + szLDAPSuffix;
	var szSQL = "select AdsPath from '" + EscapeString(szSitesLDAPPath) + "'";
	szSQL += " where Name='" + szSOMName + "'";

	// Execute the search
	ADOCommand.CommandText = szSQL;

	try
	{
		var rs = ADOCommand.Execute();
	}
	catch (err)
	{
		WScript.Echo("There was an error executing the DS query " + szSQL);
		WScript.Echo("The error was:");
		WScript.Echo(ErrCode(err.number) + " - " + err.description);
		return false;
	}

	while ( ! rs.eof )
	{
		SOM = GetObject(rs.Fields(0));
		if (SOM.Class == 'site')
		{
			GPMSOM = GPMSitesContainer.GetSite(SOM.Name)
			aResult = aResult.concat(GPMSOM);
		}

		rs.MoveNext();
	}

	// Cleanup
	ADOConnection.Close();

	// Return the result
	if (aResult.length == 1)
	{
		return aResult[0];
	}
	
	if (aResult.length == 0)
	{
		return false;
	}

	return aResult;
}

// Retrieves a specific backup from the specified location
function GetBackup(szBackupLocation, szBackupID)
{
	var GPMBackup;
	var GPMBackupDir;
	
	// Get the backup directory specified
	try
	{
		GPMBackupDir = GPM.GetBackupDir(szBackupLocation);
	}
	catch (err)
	{
		WScript.Echo("The specified backup folder '" + szBackupLocation + "' could not be accessed.");
		return false;
	}

	// See if we were passed a valid backup ID
	try
	{
		GPMBackup = GPMBackupDir.GetBackup(szBackupID);
	}
	catch (err)
	{
		GPMBackup = false;
	}
		
	if (!GPMBackup)
	{
		// Not a valid backup ID, so fetch backup by GPO name
		var GPMSearchCriteria = GPM.CreateSearchCriteria();
		GPMSearchCriteria.Add(Constants.SearchPropertyBackupMostRecent, Constants.SearchOpEquals, true);
		GPMSearchCriteria.Add(Constants.SearchPropertyGPODisplayName, Constants.SearchOpEquals, szBackupID);
		var BackupList = GPMBackupDir.SearchBackups(GPMSearchCriteria);

		if (BackupList.Count == 0)
		{
			WScript.Echo("The specified backup '" + szBackupID + "' was not found in folder '" + szBackupLocation);
			return false;
		}
		else
		{
			GPMBackup = BackupList.Item(1);
		}
	}
	
	return GPMBackup;
}

// Prints any status messages for a GPO operation, such as backup or import
function PrintStatusMessages(GPMResult)
{
	var GPMStatus = GPMResult.Status;

	if (GPMStatus.Count == 0)
	{
		// No messages, so just return
		return;
	}

	WScript.Echo("");
	var e = new Enumerator(GPMStatus);
	for (; !e.atEnd(); e.moveNext())
	{
		WScript.Echo(e.item().Message);
	}
}

// Returns the DNS domain name for the current user, using ADSI
function GetDNSDomainForCurrentUser()
{

	var ADS_NAME_INITTYPE_DOMAIN = 1;
	var ADS_NAME_INITTYPE_SERVER = 2;
	var ADS_NAME_INITTYPE_GC = 3;
 
	var ADS_NAME_TYPE_1779 = 1;                      // "CN=Jane Doe,CN=users, DC=Microsoft, DC=com"
	var ADS_NAME_TYPE_CANONICAL = 2;                 // "Microsoft.com/Users/Jane Doe".
	var ADS_NAME_TYPE_NT4 = 3;                       // "Microsoft\JaneDoe"
	var ADS_NAME_TYPE_DISPLAY = 4;                   // "Jane Doe"
	var ADS_NAME_TYPE_DOMAIN_SIMPLE = 5;             // "JaneDoe@Microsoft.com"
	var ADS_NAME_TYPE_ENTERPRISE_SIMPLE = 6;         // "JaneDoe@Microsoft.com"
	var ADS_NAME_TYPE_GUID = 7;                      // {95ee9fff-3436-11d1-b2b0-d15ae3ac8436}
	var ADS_NAME_TYPE_UNKNOWN = 8;                   // The system will try to make the best guess
	var ADS_NAME_TYPE_USER_PRINCIPAL_NAME = 9;       // "JaneDoe@Fabrikam.com"
	var ADS_NAME_TYPE_CANONICAL_EX = 10;             // "Microsoft.com/Users Jane Doe"
	var ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME = 11;   // "www/www.microsoft.com@microsoft.com"
	var ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME = 12;  // "O:AOG:DAD:(A;;RPWPCCDCLCSWRCWDWOGA;;;S-1-0-0)"
 

	var objWshNetwork = new ActiveXObject("Wscript.Network");
	var objectNameTranslate = new ActiveXObject("NameTranslate");
	var arrNamePart;
	var strNTPath = "";
	var strTranslatedName = "";
	var strResult = "";

	strUser = objWshNetwork.UserName;
	strDomain = objWshNetwork.UserDomain;
	strNTPath = strDomain + "\\" + strUser;

	objectNameTranslate.Init(ADS_NAME_INITTYPE_DOMAIN, strDomain);
	objectNameTranslate.Set(ADS_NAME_TYPE_NT4, strNTPath);
	strTranslatedName = objectNameTranslate.Get(ADS_NAME_TYPE_CANONICAL);

	arrNamePart = strTranslatedName.split("/");
	strResult = arrNamePart[0];

	return strResult;
}

// Use ADSI to get the LDAP-style forest name of a given domain
function GetForestLDAPPath(szDomainName)
{
	// Get the RootDSE naming context for the specified domain
	var RootDSE = GetObject("LDAP://" + szDomainName + "/RootDSE");

	// Initialize the property cache
	RootDSE.GetInfo();

	// Now get the forest name
	var szForestName = RootDSE.rootDomainNamingContext;
	
	return szForestName;
}

// Use ADSI to get the forest name of a given domain
function GetForestDNSName(szDomainName)
{
	var ADS_NAME_INITTYPE_DOMAIN = 1;
	var ADS_NAME_INITTYPE_SERVER = 2;
	var ADS_NAME_INITTYPE_GC = 3;
 
	var ADS_NAME_TYPE_1779 = 1;                      // "CN=Jane Doe,CN=users, DC=Microsoft, DC=com"
	var ADS_NAME_TYPE_CANONICAL = 2;                 // "Microsoft.com/Users/Jane Doe".
	var ADS_NAME_TYPE_NT4 = 3;                       // "Microsoft\JaneDoe"
	var ADS_NAME_TYPE_DISPLAY = 4;                   // "Jane Doe"
	var ADS_NAME_TYPE_DOMAIN_SIMPLE = 5;             // "JaneDoe@Microsoft.com"
	var ADS_NAME_TYPE_ENTERPRISE_SIMPLE = 6;         // "JaneDoe@Microsoft.com"
	var ADS_NAME_TYPE_GUID = 7;                      // {95ee9fff-3436-11d1-b2b0-d15ae3ac8436}
	var ADS_NAME_TYPE_UNKNOWN = 8;                   // The system will try to make the best guess
	var ADS_NAME_TYPE_USER_PRINCIPAL_NAME = 9;       // "JaneDoe@Fabrikam.com"
	var ADS_NAME_TYPE_CANONICAL_EX = 10;             // "Microsoft.com/Users Jane Doe"
	var ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME = 11;   // "www/www.microsoft.com@microsoft.com"
	var ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME = 12;  // "O:AOG:DAD:(A;;RPWPCCDCLCSWRCWDWOGA;;;S-1-0-0)"


	// Get the RootDSE naming context for the specified domain
	var RootDSE = GetObject("LDAP://" + szDomainName + "/RootDSE");

	// Initialize the property cache
	RootDSE.GetInfo();

	// Now get the forest name
	var szForestName = RootDSE.rootDomainNamingContext;

	// Translate it to DNS style
	var objectNameTranslate = new ActiveXObject("NameTranslate");
	objectNameTranslate.Init(ADS_NAME_INITTYPE_DOMAIN, szDomainName);
	objectNameTranslate.Set(ADS_NAME_TYPE_1779, szForestName);

	var szTranslatedName = objectNameTranslate.Get(ADS_NAME_TYPE_CANONICAL);
	
	return szTranslatedName.slice(0,-1);
}

// Escapes certain characters in a string so they will work with SQL statements
function EscapeString(str)
{
	var result;

	// Handle single quotes
	var re = new RegExp(/'/g);
	result = str.replace(re, "''");
	return result;
}

// Replaces invalid characters in a file name
function GetValidFileName(str)
{
    var result = str;
    result = result.replace(/\*/g, "");
    result = result.replace(/\\/g, "");
    result = result.replace(/\//g, "");
    result = result.replace(/\|/g, "");
    result = result.replace(/>/g, "");
    result = result.replace(/</g, "");
    result = result.replace(/:/g, "");
    result = result.replace(/\"/g, "");
    result = result.replace(/\?/g, "");

    return result;
}

// Checks if the specified file system path is valid.
// Returns true if the path is found, false otherwise.
//
function ValidatePath(szPath)
{
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	try
	{
		var Path = fso.GetFolder(szPath);
	}
	catch (err)
	{
		return false;
	}
	
	return true;
}

// Returns the hexadecimal string for a number, converting negative decimal
// values to the appropriate winerror style hex values
//
function ErrCode(i)
{
	var result;

	if (i < 0)
	{
		// Get the winerror-style representation of the hex value
		result = 0xFFFFFFFF + i + 1;
	}
	else
	{
		result = i;
	}

	return "0x" + result.toString(16); // base 16
}
