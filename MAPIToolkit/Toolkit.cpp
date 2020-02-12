#pragma once
// MAPIToolkit.cpp : Defines the functions for the static library.
//

#include "framework.h"
#include <mapidefs.h>
#include <guiddef.h>
#include <initguid.h>
#define USES_IID_IMAPIProp
#define USES_IID_IMsgServiceAdmin2
#include <atlchecked.h>
#include "ToolkitTypeDefs.h"
#include "RegistryHelper.h"
#include <iterator> 
#include <algorithm>
#include "Logger.h"
#include "Toolkit.h"
#include "InlineAndMacros.h"
#include "Profile//Profile.h"
#include "Misc/Utility/StringOperations.h"
#include "Misc/XML/AddressBookXmlParser.h"
#include <vector>


#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "Mapi32.lib")
#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shlwapi.lib")

#pragma warning(disable:4996) // _CRT_SECURE_NO_WARNINGS

namespace MAPIToolkit
{
	std::map<std::wstring, ULONG> Toolkit::g_actionsMap =
	{
		{ L"addprofile",			ACTION_PROFILE_ADD},
		{ L"addprovider",			ACTION_PROVIDER_ADD },
		{ L"addservice",			ACTION_SERVICE_ADD },
		{ L"changedatafilepath",	ACTION_SERVICE_CHANGEDATAFILEPATH },
		{ L"cloneprofile",			ACTION_PROFILE_CLONE },
		{ L"promotedelegates",		ACTION_PROFILE_PROMOTEDELEGATES },
		{ L"listallprofiles",		ACTION_PROFILE_LISTALL },
		{ L"listallproviders",		ACTION_PROVIDER_LISTALL },
		{ L"listallservices",		ACTION_SERVICE_LISTALL },
		{ L"listprofile",			ACTION_PROFILE_LIST },
		{ L"listprovider",			ACTION_PROVIDER_LIST },
		{ L"listservice",			ACTION_SERVICE_LIST },
		{ L"promoteonedelegate",	ACTION_PROFILE_PROMOTEONEDELEGATE },
		{ L"removeallprofiles",		ACTION_PROFILE_REMOVEALL },
		{ L"removeallproviders",	ACTION_PROVIDER_REMOVEALL },
		{ L"removeallservices",		ACTION_SERVICE_REMOVEALL },
		{ L"removeprofile",			ACTION_PROFILE_REMOVE },
		{ L"removeprovider",		ACTION_PROVIDER_REMOVE },
		{ L"removeservice",			ACTION_SERVICE_REMOVE },
		{ L"setcachedmode",			ACTION_SERVICE_SETCACHEDMODE },
		{ L"setdefaultprofile",		ACTION_PROFILE_SETDEFAULT },
		{ L"setdefaultservice",		ACTION_SERVICE_SETDEFAULT },
		{ L"renameprofile",			ACTION_PROFILE_RENAME },
		{ L"updateprovider",		ACTION_PROVIDER_UPDATE },
		{ L"updateservice",			ACTION_SERVICE_UPDATE }
	};

	std::map<std::wstring, ULONG> Toolkit::g_profileModeMap =
	{
		{ L"default",	PROFILEMODE_DEFAULT},
		{ L"specific",	PROFILEMODE_SPECIFIC },
		{ L"all",		PROFILEMODE_ALL }
	};

	std::map<std::wstring, ULONG> Toolkit::g_serviceModeMap =
	{
		{ L"default",	SERVICEMODE_DEFAULT},
		{ L"specific",	SERVICEMODE_SPECIFIC },
		{ L"all",		SERVICEMODE_ALL }
	};

	std::map<std::wstring, ULONG> Toolkit::g_saveConfigMap =
	{
		{ L"true",	SAVECONFIG_TRUE},
		{ L"false",	SAVECONFIG_FALSE }
	};

	std::map<std::wstring, ULONG> Toolkit::g_serviceTypeMap =
	{
		{ L"addressbook",		SERVICETYPE_ADDRESSBOOK},
		{ L"datafile",			SERVICETYPE_DATAFILE },
		{ L"exchangeaccount",	SERVICETYPE_EXCHANGEACCOUNT },
		{ L"all",				SERVICETYPE_ALL }
	};

	std::map<std::wstring, std::wstring> Toolkit::g_addressBookMap =
	{
		{ L"servicename",		L"EMABLT"},
		{ L"displayname",		L""},
		{ L"servername",		L"" },
		{ L"serverport",		L"389" },
		{ L"usessl",			L"false" },
		{ L"username",			L"" },
		{ L"password",			L"" },
		{ L"requirespa",		L"false" },
		{ L"searchtimeout",		L"60" },
		{ L"maxentries",		L"100" },
		{ L"defaultsearchbase",	L"true" },
		{ L"customsearchbase",	L"" },
		{ L"enablebrowsing",	L"false" },
		{ L"newservername",		L"" },
		{ L"newdisplayname",	L"" },
		{ L"newserverport",		L"" }
	};


#define LOGGINGMODE_UNKNOWN			0x00000000
#define LOGGINGMODE_NONE			0x00000001
#define LOGGINGMODE_CONSOLE			0x00000002
#define LOGGINGMODE_FILE			0x00000004
#define LOGGINGMODE_ALL				0x00000008
#define LOGGINGMODE_VERBOSE			0x00000010

	std::map<std::wstring, ULONG> Toolkit::g_loggingModeMap =
	{
		{ L"none",		LOGGINGMODE_NONE },
		{ L"console",	LOGGINGMODE_CONSOLE },
		{ L"file",		LOGGINGMODE_FILE },
		{ L"all",		LOGGINGMODE_ALL },
		{ L"debug",		LOGGINGMODE_DEBUG }
	};

	std::map<std::wstring, std::wstring> Toolkit::g_toolkitMap =
	{
		{ L"action",			L""},
		{ L"outlookversion",	L"" },
		{ L"loggingmode",		L"console" },
		{ L"exportpath",		L"" },
		{ L"exportmode",		L"" },
		{ L"logfilepath",		L"" },
		{ L"profilemode",		L"default" },
		{ L"profilename",		L""},
		{ L"profilecount",		L""},
		{ L"servicemode",		L"default" },
		{ L"servicetype",		L"" },
		{ L"providermode",		L"default" },
		{ L"providertype",		L"" },
		{ L"configfilepath",	L"" },
		{ L"saveconfig",		L"false"},
		{ L"setfirstab",		L"false"},
		{ L"currentprofilename",L""}
	};

	std::map<std::wstring, std::wstring> Toolkit::g_regKeyMap =
	{
		{ L"toolkit",				L"SOFTWARE\\Microsoft Ltd\\MAPIToolkit"},
		{ L"addressbook",			L"SOFTWARE\\Microsoft Ltd\\MAPIToolkit\\AddressBook" }
	};

	std::map <int, std::wstring > Toolkit::g_hexMap =
	{
		{ 0, L"0"},
		{ 1, L"1"},
		{ 2, L"2"},
		{ 3, L"3"},
		{ 4, L"4"},
		{ 5, L"5"},
		{ 6, L"6"},
		{ 7, L"7"},
		{ 8, L"8"},
		{ 9, L"9"},
		{ 10, L"A"},
		{ 11, L"B"},
		{ 12, L"C"},
		{ 13, L"D"},
		{ 14, L"E"},
		{ 15, L"F"}
	};

	std::map <std::wstring, std::wstring > Toolkit::g_parameterHelpInfo =
	{
		{ L"?",					L"Displays the help info."},
		{ L"registry",			L"Indicates whether to read the configuration from the registry if previously saved with \"-saveconfig true\"."},
		{ L"action",			L"Action(s) to perform."},
		{ L"action",			L"Action(s) to perform."},
		{ L"configfilepath",	L"Path to the input configuration file." },
		{ L"customsearchbase",	L"custom search base in case defaultsearchbase is set to false." },
		{ L"defaultsearchbase",	L"If \"true\" the default search base is to be used. The default value is 'true'." },
		{ L"displayname",		L"The display name of the service to run the action(s) against."},
		{ L"enablebrowsing",	L"Indicates whether browsing the address book contens is supported. " },
		{ L"logfilepath",		L"Path towards the log file where informatin is to be logged." },
		{ L"loggingmode",		L"Indicates how logging is captured." },
		{ L"maxentries",		L"The maximum number of results returned by a search in the address book. The default value is 100." },
		{ L"newdisplayname",	L"Display name to replace the current display name of the service with." },
		{ L"newservername",		L"Server name to replace the current server name with in the speciifed service." },
		{ L"newserverport",		L"Server port to replace the current server port with in the speciifed service." },
		{ L"password",			L"The password to use for authenticating. This must be a clear text passord. It will be encrypted via CryptoAPI and stored in the address book settings." },
		{ L"profilemode",		L"Indicates whether to run the action on all profiles or a specific profile." },
		{ L"profilename",		L"Indicates the name of the profile to run the action against. If left empty, the default profile will be used, unles the profilemode specified is \"all\"."},
		{ L"requirespa",		L"\"true\" if Secure Password Authentication is required is required. The default value is \"false\"" },
		{ L"saveconfig",		L"Indicates whether to save the current configuration in teh registry or no"},
		{ L"searchtimeout",		L"The number of seconds before the search request times out. The default value is 60 seconds." },
		{ L"servername",		L"The LDAP address book server address. For example \"ldap.contoso.com\"." },
		{ L"serverport",		L"The LDAP port to connect to. The standard port for Active Directory is 389." },
		{ L"servicetype",		L"Indicates the type of service to run the action against." },
		{ L"setfirstab",		L"Sets the newly created Address Book objet as first Address Book in the Address book search order in Outlook." },
		{ L"username",			L"The Username to use for authenticating in the form of domain\\username, UPN or just the username if domain name not applicable or not required. Leave blank if a username and password are not required." },
		{ L"usessl",			L"\"true\" if a SSL connection is required.The default value is \"false\"." }
	};

	std::map <std::wstring, std::wstring > Toolkit::g_parameterHelpValues =
	{
		{ L"action",			L"{addservice, listservice, listallservices, removeservice, removeallservices, updateservice}"},
		{ L"configfilepath",	L"<string>" },
		{ L"customsearchbase",	L"<string>" },
		{ L"defaultsearchbase",	L"{true, false}" },
		{ L"displayname",		L"<string>"},
		{ L"enablebrowsing",	L"{true, false}" },
		{ L"exportpath",		L"<string>" },
		{ L"logfilepath",		L"<string>" },
		{ L"loggingmode",		L"{none, console, file, all, debug}" },
		{ L"maxentries",		L"<int>" },
		{ L"newdisplayname",	L"<string>" },
		{ L"newservername",		L"<string>" },
		{ L"newserverport",		L"<int>" },
		{ L"password",			L"<string>" },
		{ L"profilemode",		L"{default, specific, all}" },
		{ L"profilename",		L"<string>"},
		{ L"requirespa",		L"{true, false}" },
		{ L"saveconfig",		L"{true, false}"},
		{ L"searchtimeout",		L"<int>" },
		{ L"servername",		L"<string>" },
		{ L"serverport",		L"<int>" },
		{ L"servicetype",		L"{addressbook}" },
		{ L"setfirstab",		L"{{true, false}}" },
		{ L"username",			L"<string>" },
		{ L"usessl",			L"{true, false}" }
	};

	ULONG Toolkit::m_action;
	LPPROFADMIN Toolkit::m_pProfAdmin;
	BOOL Toolkit::m_registry = FALSE;
	LPMAPISESSION Toolkit::m_lpMapiSession = NULL;
	BOOL Toolkit::m_bLoggedOn = FALSE;

	// Is64BitProcess
// Returns true if 64 bit process or false if 32 bit.
	BOOL Toolkit::Is64BitProcess(void)
	{
#if defined(_WIN64)
		return TRUE;   // 64-bit program
#else
		return FALSE;
#endif
	}

	void Toolkit::DisplayUsage()
	{
		std::wprintf(L"MAPIToolkit - MAPI profile utility\n");
		std::wprintf(L"       Allows the management of Outlook / MAPI profiles at the command line.\n");
		std::wprintf(L"Usage:\n");
		for (auto const& keyValuePair : g_parameterHelpInfo)
		{
			try
			{
				if (!g_parameterHelpValues.at(keyValuePair.first).empty())
				{
					std::wstring wszLine = L"       [-" + SpaceIt(keyValuePair.first, 20) + L" " + g_parameterHelpValues.at(keyValuePair.first) + L"]\n";
					std::wprintf(wszLine.c_str());
				}
			}
			catch (const std::exception& e)
			{
				std::wstring wszLine = L"       [-" + keyValuePair.first+ L"]\n";
				std::wprintf(wszLine.c_str());
			}
		}

		std::wprintf(L"Options:\n");

		for (auto const& keyValuePair : g_parameterHelpInfo)
		{
			std::wstring wszLine = L" -" + SpaceIt(keyValuePair.first, 20) + L": " + keyValuePair.second + L"\n";
				std::wprintf(wszLine.c_str());
		}

	/*	std::wprintf(L"MAPIToolkit - MAPI profile utility\n");
		std::wprintf(L"       Allows the management of Outlook / MAPI profiles at the command line.\n");
		std::wprintf(L"\n");
		std::wprintf(L"Usage: [-?] \n");
		std::wprintf(L"       [-action <addservice, listallservices, listservice, removeallservices, removeservice, updateservice>]\n");
		std::wprintf(L"       [-profilemode <default, specific, all>]\n");
		std::wprintf(L"       [-profilename name]\n");
		std::wprintf(L"       [-servicetype <addressbook>]\n");
		std::wprintf(L"       [-displayname name]\n");
		std::wprintf(L"       [-servername name]\n");
		std::wprintf(L"       [-serverport port]\n");
		std::wprintf(L"       [-usessl <true, false>]\n");
		std::wprintf(L"       [-username username]\n");
		std::wprintf(L"       [-password password]\n");
		std::wprintf(L"       [-requirespa <true, false>]\n");
		std::wprintf(L"       [-searchtimeout timeout]\n");
		std::wprintf(L"       [-maxentries maxentries]\n");
		std::wprintf(L"       [-defaultsearchbase <true, false>]\n");
		std::wprintf(L"       [-customsearchbase searchbase]\n");
		std::wprintf(L"       [-enablebrowsing <true, false>]\n");
		std::wprintf(L"       [-configfilepath path]\n");
		std::wprintf(L"       \n");
		std::wprintf(L"Options:\n");
		std::wprintf(L" -?:			         Displays this information.\n");
		std::wprintf(L" -action:             \"addservice\" adds a service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                      \"listallservices\" lists all services with the type specified by \"servicetype\".\n");
		std::wprintf(L"                      \"listservice\" lists a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                      \"removeallservices\" removes all services with the type specified by \"servicetype\".\n");
		std::wprintf(L"                      \"removeservice\" removes a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L"                      \"updateservice\" updates a specific service with the type specified by \"servicetype\".\n");
		std::wprintf(L" -profilemode:        \"default\" to run the selected action on the default profile.\n");
		std::wprintf(L"                      \"specific\" to run the selected action on the profile specified by the \"profilename\" value.\n");
		std::wprintf(L"                      \"all\" to run the selected action against all profiles.\n");
		std::wprintf(L"                      The default profile will be used if a profile mode is not specified.\n");
		std::wprintf(L" -profilename:        Name of the profile to run the specified actiona against. The profile name is mandatory\n");
		std::wprintf(L"                      if \"profilename\" is set to \"specific\" \n");
		std::wprintf(L" -servicetype:        \"addressbook\" to run an addressbook specific operation.\n");
		std::wprintf(L" -servicetype:        This is the only operation currently supported.\n");
		std::wprintf(L" -displayname:        The display name of the address book service to create, update, list or remove.\n");
		std::wprintf(L" -servername:         The display name of the LDAP server configured in the address book.\n");
		std::wprintf(L" -configfilepath:     The path towards the address book configuration XML to use.\n");
		std::wprintf(L" -serverport:         The LDAP port to connect to. The standard port for Active Directory is 389.\n");
		std::wprintf(L" -usessl:             \"true\" if a SSL connection is required. The default value is \"false\".\n");
		std::wprintf(L" -username:           The Username to use for authenticating in the form of domain\\username, UPN or just the username \n");
		std::wprintf(L"                      if domain name not applicable or not required. Leave blank if a username and password are \n");
		std::wprintf(L"                      not required.\n");
		std::wprintf(L" -password:           The Password to use for authenticating. This must be a clear text passord. It will be encrypted \n");
		std::wprintf(L"                      via CryptoAPIand stored in the AB settings. \n");
		std::wprintf(L" -requirespa:         \"true\" if Secure Password Authentication is required is required. The default value is \"false\".\n");
		std::wprintf(L" -searchtimeout:      The number of seconds before the search request times out. The default value is 60 seconds.]\n");
		std::wprintf(L" -maxentries:         The maximum number of results returned by a search in this AB. The default value is 100.\n");
		std::wprintf(L" -defaultsearchbase:  \"true\" the default search base is to be used. The default value is \"true\". \n");
		std::wprintf(L" -customsearchbase:   Custom search base in case DefaultSearchBase is set to False. \n");
		std::wprintf(L" -enablebrowsing:     Indicates whether browsing the AB contens is supported.\n");
		std::wprintf(L" -configfilepath:     The path towards the address book configuration XML to use.\n");*/
	}

	// GetOutlookVersion
	int Toolkit::GetOutlookVersion()
	{
		std::wstring szOLVer = L"";
		szOLVer = GetRegStringValue(HKEY_CLASSES_ROOT, TEXT("Outlook.Application\\CurVer"), NULL);
		if (szOLVer != L"")
		{
			if (szOLVer == L"Outlook.Application.16")
			{
				return 2016;
			}
			else if (szOLVer == L"Outlook.Application.15")
			{
				return 2013;
			}
			else if (szOLVer == L"Outlook.Application.14")
			{
				return 2010;
			}
			else if (szOLVer == L"Outlook.Application.12")
			{
				return 2007;
			}
			return 0;
		}
		else return 0;
	}

	// IsCorrectBitness
	// Matches the App bitness against Outlook's bitness to avoid MAPI version errors at startup
	// The execution will only continue if the bitness is matched.
	BOOL _cdecl Toolkit::IsCorrectBitness()
	{
		std::wstring szOLVer = L"";
		std::wstring szOLBitness = L"";
		szOLVer = GetRegStringValue(HKEY_CLASSES_ROOT, TEXT("Outlook.Application\\CurVer"), NULL);
		if (szOLVer != L"")
		{
			if (szOLVer == L"Outlook.Application.19")
			{
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\16.0\\Outlook"), TEXT("Bitness"));
				if (szOLBitness != L"")
				{
					if (szOLBitness == L"x64")
					{
						if (Toolkit::Is64BitProcess())
							return TRUE;
					}
					else if (szOLBitness == L"x86")
					{
						if (Is64BitProcess())
							return FALSE;
						else
							return TRUE;
					}
					else return FALSE;
				}
			}
			else if (szOLVer == L"Outlook.Application.16")
			{
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\16.0\\Outlook"), TEXT("Bitness"));
				if (szOLBitness != L"")
				{
					if (szOLBitness == L"x64")
					{
						if (Is64BitProcess())
							return TRUE;
					}
					else if (szOLBitness == L"x86")
					{
						if (Is64BitProcess())
							return FALSE;
						else
							return TRUE;
					}
					else return FALSE;
				}
			}
			else if (szOLVer == L"Outlook.Application.15")
			{
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\15.0\\Outlook"), TEXT("Bitness"));
				if (szOLBitness != L"")
				{
					if (szOLBitness == L"x64")
					{
						if (Is64BitProcess())
							return TRUE;
					}
					else if (szOLBitness == L"x86")
					{
						if (Is64BitProcess())
							return FALSE;
						else
							return TRUE;
					}
					else return FALSE;
				}
			}
			else if (szOLVer == L"Outlook.Application.14")
			{
				szOLBitness = GetRegStringValue(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Office\\14.0\\Outlook"), TEXT("Bitness"));
				if (szOLBitness != L"")
				{
					if (szOLBitness == L"x64")
					{
						if (Is64BitProcess())
							return TRUE;
					}
					else if (szOLBitness == L"x86")
					{
						if (Is64BitProcess())
							return FALSE;
						else
							return TRUE;
					}
					else return FALSE;
				}
			}
			else if (szOLVer == L"Outlook.Application.12")
			{
				if (Is64BitProcess())
					return FALSE;
			}
			else return FALSE;
			return FALSE;
		}
		else return FALSE;
	}

	BOOL Toolkit::Initialise()
	{
		HRESULT hRes = S_OK;
		m_action = ACTION_UNSPECIFIED;
		m_pProfAdmin = NULL;
		MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
		// parse the actions
		std::wstring wszActionItem;
		std::wstringstream wss;
		CHK_HR_DBG(CoInitialize(NULL), L"CoInitialize");
		if (!g_toolkitMap.at(L"logfilepath").empty())
			Logger::SetFilePath(g_toolkitMap.at(L"logfilepath"));
		if (!m_registry)
			if (!g_toolkitMap.at(L"configfilepath").empty())
				CHK_HR_DBG(ParseAddressBookXml((LPTSTR)g_toolkitMap.at(L"configfilepath").c_str()), L"ParseAddressBookXml");

		wss << g_toolkitMap.at(L"action");
		while (std::getline(wss, wszActionItem, L'|'))
		{
			try
			{
				m_action |= g_actionsMap.at(wszActionItem);
			}
			catch (const std::exception& e)
			{

			}
		}


		CHK_HR_DBG(MAPIInitialize(&MAPIINIT), L"MAPIInitialize");

		CHK_HR_DBG(MAPIAdminProfiles(0, &m_pProfAdmin), L"MAPIAdminProfiles");
		if (VCHK(g_profileModeMap.at(g_toolkitMap.at(L"profilemode")), PROFILEMODE_DEFAULT))
		{
			g_toolkitMap.at(L"profilename") = GetDefaultProfileName(m_pProfAdmin);
			g_toolkitMap.at(L"currentprofilename") = g_toolkitMap.at(L"profilename");
			g_toolkitMap.at(L"profilemode") = L"specific";
		}
		g_toolkitMap.at(L"profilecount") = ConvertIntToString(GetProfileCount(m_pProfAdmin));

		MAPIAllocateBuffer(sizeof(LPMAPISESSION), (LPVOID*)&m_lpMapiSession);
		ZeroMemory(m_lpMapiSession, sizeof(LPMAPISESSION));

	Error:
		goto CleanUp;
	CleanUp:
		return TRUE;
	}

	VOID Toolkit::Uninitialise()
	{
		if (m_pProfAdmin) m_pProfAdmin->Release();
		if (m_lpMapiSession) HrLogoff(m_lpMapiSession);
		MAPIUninitialize();
		CoUninitialize();
	}

	BOOL Toolkit::SaveConfig()
	{
		for (auto const& keyValuePair : g_toolkitMap)
		{
			if (!keyValuePair.second.empty())
				if (!WriteRegStringValue(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"toolkit").c_str(), (LPCTSTR)keyValuePair.first.c_str(), (LPCTSTR)keyValuePair.second.c_str())) return FALSE;
		}

		if (0 == wcscmp(g_toolkitMap.at(L"servicetype").c_str(), L"addressbook"))
		{
			for (auto const& keyValuePair : g_addressBookMap)
			{
				if (!keyValuePair.second.empty())
					if (!WriteRegStringValue(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"addressbook").c_str(), (LPCTSTR)keyValuePair.first.c_str(), (LPCTSTR)keyValuePair.second.c_str())) return FALSE;
			}
		}
		return TRUE;
	}

	BOOL Toolkit::ReadConfig()
	{
		ReadRegConfig(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"toolkit").c_str(), &g_toolkitMap);
		ReadRegConfig(HKEY_CURRENT_USER, (LPCTSTR)g_regKeyMap.at(L"addressbook").c_str(), &g_addressBookMap);
		return TRUE;
	}

	void Toolkit::Run(int argc, wchar_t* argv[])
	{
		if (ParseParams(argc, argv))
		{
			Logger::SetLoggingMode(g_loggingModeMap.at(g_toolkitMap.at(L"loggingmode")));
			Initialise();
			RunAction();
			if (SAVECONFIG_TRUE == g_saveConfigMap.at(g_toolkitMap.at(L"saveconfig")))
				SaveConfig();
		}
		else
			DisplayUsage();

		Uninitialise();
	}

	VOID Toolkit::CustomRun()
	{
		HRESULT hRes = S_OK;
		m_pProfAdmin = NULL;
		LPSRowSet     pNewRows = NULL;
		MAPIINIT_0  MAPIINIT = { 0, MAPI_MULTITHREAD_NOTIFICATIONS };
		CHK_HR_DBG(CoInitialize(NULL), L"CoInitialize");
		CHK_HR_DBG(MAPIInitialize(&MAPIINIT), L"MAPIInitialize");
		CHK_HR_DBG(MAPIAdminProfiles(0, &m_pProfAdmin), L"MAPIAdminProfiles");
		g_toolkitMap.at(L"profilename") = GetDefaultProfileName(m_pProfAdmin);
		g_toolkitMap.at(L"currentprofilename") = g_toolkitMap.at(L"profilename");
		g_toolkitMap.at(L"profilemode") = L"specific";

		MAPIAllocateBuffer(sizeof(LPMAPISESSION), (LPVOID*)& m_lpMapiSession);
		ZeroMemory(m_lpMapiSession, sizeof(LPMAPISESSION));

		CHK_HR_DBG(MAPILogonEx(NULL, (LPTSTR)g_toolkitMap.at(L"currentprofilename").c_str(), NULL, 0, &m_lpMapiSession), L"MAPILogonEx");
		
		
		MAPIAllocateBuffer(sizeof(LPSRowSet), (LPVOID*)& pNewRows);
		ZeroMemory(pNewRows, sizeof(LPSRowSet));
		HrGetABSearchOrder(m_lpMapiSession, &pNewRows);

	Error:
		goto Cleanup;
	Cleanup:
		if (m_pProfAdmin) m_pProfAdmin->Release();
		if (m_lpMapiSession) HrLogoff(m_lpMapiSession);
		MAPIUninitialize();
		CoUninitialize();
	}

	VOID Toolkit::AddService(LPSERVICEADMIN2 pServiceAdmin)
	{
		switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
		{
		case SERVICETYPE_ADDRESSBOOK:
		{
			if (AddAddressBookService(pServiceAdmin))
				Logger::WriteLine(LOGLEVEL_SUCCESS, L"Address book service succesfully added");

			break;
		}
		case SERVICETYPE_EXCHANGEACCOUNT:
		case SERVICETYPE_DATAFILE:
		{
			Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		}
	}

	BOOL Toolkit::AddAddressBookService(LPSERVICEADMIN2 pServiceAdmin2)
	{
		HRESULT hRes = S_OK;
		if (!g_addressBookMap.at(L"servername").empty())
		{
			// if no display name specified, use the server name instead
			if (g_addressBookMap.at(L"displayname").empty())
			{
				g_addressBookMap.at(L"displayname") = g_addressBookMap.at(L"servername");
			}
			BOOL bServiceExists = false;
			ULONG ulCServices;

			// Get service UID(s) for the services we want to remove	
			ULONG cServices = 0;
			if SUCCEEDED(GetABServiceUid(pServiceAdmin2, g_addressBookMap.at(L"displayname").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), g_addressBookMap.at(L"servername").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"servername").c_str(), &cServices, NULL))
			{
				if (cServices == 0)
				{
					CHK_HR_DBG(CreateABService(pServiceAdmin2), L"CreateABService");
					if (!g_toolkitMap.at(L"setfirstab").empty())
					{
						CHK_HR_DBG(HrSetABSearchOrder((WCHAR*)g_addressBookMap.at(L"displayname").c_str(),ConvertStdStringToWideChar(g_addressBookMap.at(L"currentprofilename").c_str())), L"HrSetABSearchOrder");
					}
				}
				else
				{
					Logger::WriteLine(LOGLEVEL_FAILED, L"A service with the requested configuration already exists");
					return FALSE;
				}
			}
			else
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"An address book service with the selected configuration already exists");
				return FALSE;
			}
		}
	Error:
		goto CleanUp;
	CleanUp:
		return (SUCCEEDED(hRes));
	}

	void Toolkit::UpdateService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID lpMAPIUid)
	{
		HRESULT hRes = S_OK;
		switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
		{
		case SERVICETYPE_ADDRESSBOOK:
		{
			CHK_HR_DBG(UpdateABService(pServiceAdmin, lpMAPIUid), L"UpdateABService");
			g_addressBookMap.at(L"newdisplayname") = L"";
			g_addressBookMap.at(L"newservername") = L"";
			g_addressBookMap.at(L"newserverport") = L"";
			break;
		}
		case SERVICETYPE_EXCHANGEACCOUNT:
		case SERVICETYPE_DATAFILE:
		{
			Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		}
	Error:
		goto CleanUp;
	CleanUp:
		return;
	}

	VOID Toolkit::UpdateAddressBookService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID lpMAPIUid)
	{
		HRESULT hRes = S_OK;
		if (g_addressBookMap.at(L"servername").empty() && g_addressBookMap.at(L"displayname").empty())
		{
			Logger::WriteLine(LOGLEVEL_FAILED, L"You must specify a -displayname, a -servername or both");
		}

		// Get service UID(s) for the services we want to remove	
		ULONG cServices = 0;
		if SUCCEEDED(GetABServiceUid(pServiceAdmin, &cServices, NULL), L"Getting service count");
		if (cServices > 0)
		{

			MAPIUID* pMAPIUid = new MAPIUID[cServices];
			ZeroMemory(pMAPIUid, sizeof(MAPIUID) * cServices);

			if SUCCEEDED(GetABServiceUid(pServiceAdmin, &cServices, pMAPIUid), L"Fetching existing service UIDs")
				if ((cServices > 0) && pMAPIUid)
				{
					Logger::WriteLine(LOGLEVEL_INFO, L"Action " + g_toolkitMap.at(L"action") + L" will run against " + ConvertIntToString(cServices) + L" services");
					for (int i = 0; i < cServices; i++)
					{
						Logger::WriteLine(LOGLEVEL_INFO, L"Updating service #" + ConvertIntToString(i));
						if SUCCEEDED(UpdateABService(pServiceAdmin, &pMAPIUid[i]))
							Logger::WriteLine(LOGLEVEL_SUCCESS, L"Address book service succesfully updated");
					}
				}
		}

	Error:
		goto CleanUp;
	CleanUp:
		return;
	}

	VOID Toolkit::ListService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID lpMAPIUid)
	{
		HRESULT hRes = S_OK;
		switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
		{
			case SERVICETYPE_ADDRESSBOOK:
			{
				CHK_HR_DBG(ListABService(pServiceAdmin, lpMAPIUid), L"ListABService");
				break;
			}
			case SERVICETYPE_EXCHANGEACCOUNT:
			case SERVICETYPE_DATAFILE:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
		}
	Error:
		goto CleanUp;
	CleanUp:
		return;
	}

	VOID Toolkit::RemoveAllServices(LPSERVICEADMIN2 pServiceAdmin)
	{
		switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
		{
		case SERVICETYPE_ADDRESSBOOK:
		{
			// Get service UID(s) for the services we want to remove	
			ULONG cServices = 0;
			if SUCCEEDED(GetABServiceUid(pServiceAdmin, &cServices, NULL), L"Getting service count");
			if (cServices > 0)
			{

				MAPIUID* pMAPIUid = new MAPIUID[cServices];
				ZeroMemory(pMAPIUid, sizeof(MAPIUID) * cServices);

				if SUCCEEDED(GetABServiceUid(pServiceAdmin, &cServices, pMAPIUid), L"Fetching existing service UIDs")
					if ((cServices > 0) && pMAPIUid)
					{	
						Logger::WriteLine(LOGLEVEL_INFO, L"Action " + g_toolkitMap.at(L"action") + L" will run against " + ConvertIntToString(cServices) + L" services");
						Logger::WriteLine(LOGLEVEL_INFO, L"Number of services found: " + ConvertIntToString(cServices));
						for (int i = 0; i < cServices; i++)
						{
							RemoveABService(pServiceAdmin, &pMAPIUid[i]);
						}
					}
			}
			break;
		}
		case SERVICETYPE_EXCHANGEACCOUNT:
		case SERVICETYPE_DATAFILE:
		{
			Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		}
	}

	BOOL Toolkit::GetLoggedOn()
	{
		return m_bLoggedOn;
	}

	VOID Toolkit::SetLoggedOn(BOOL bLoggedOn)
	{
		m_bLoggedOn = bLoggedOn;
	}

	void Toolkit::RunAction()
	{
		HRESULT hRes = S_OK;
		switch (m_action)
		{

		case ACTION_UNSPECIFIED:
		{
			Logger::WriteLine(LOGLEVEL_FAILED, L"You must specify an action");
			break;
		}
		case ACTION_PROFILE_ADD:
		case ACTION_PROFILE_CLONE:
		case ACTION_PROFILE_RENAME:
		case ACTION_PROFILE_LIST:
		case ACTION_PROFILE_LISTALL:
		case ACTION_PROFILE_REMOVE:
		case ACTION_PROFILE_REMOVEALL:
		case ACTION_PROFILE_SETDEFAULT:
		case ACTION_PROFILE_PROMOTEDELEGATES:
		case ACTION_PROFILE_PROMOTEONEDELEGATE:
		{
			Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
			break;
		}
		default:
		{
			std::vector<std::wstring> vProfileNames;

			switch (g_profileModeMap.at(g_toolkitMap.at(L"profilemode")))
			{
			case PROFILEMODE_ALL:

				CHK_BOOL_MSG(GetProfileNames(m_pProfAdmin, &vProfileNames), L"Retrieving profile names");
				for (auto const& profileName : vProfileNames) {
					g_toolkitMap.at(L"currentprofilename") = profileName;
					//CHK_HR_DBG(HrLogon((LPTSTR)g_toolkitMap.at(L"currentprofilename").c_str(), &m_lpMapiSession), L"HrLogon");
					if (RunActionOneProfile(profileName))
					{
						Logger::WriteLine(LOGLEVEL_SUCCESS, L"Action succesfully run on profile: " + profileName);
					}
					else
					{
						Logger::WriteLine(LOGLEVEL_FAILED, L"Unable to run action on profile: " + profileName);
					}
				}
				break;
			case PROFILEMODE_SPECIFIC:
				//CHK_HR_DBG(HrLogon((LPTSTR)g_toolkitMap.at(L"profilename").c_str(), &m_lpMapiSession), L"HrLogon");
				if (RunActionOneProfile(g_toolkitMap.at(L"profilename")))
				{
					Logger::WriteLine(LOGLEVEL_SUCCESS, L"Action succesfully run on profile: " + g_toolkitMap.at(L"profilename"));
				}
				else
				{
					Logger::WriteLine(LOGLEVEL_FAILED, L"Unable to run action on profile: " + g_toolkitMap.at(L"profilename"));
				}
				break;
			}
			break;
		}

		}

	Error:
		goto CleanUp;
	CleanUp:
		return;
	}

	BOOL Toolkit::RunActionOneProfile(std::wstring wszProfileName)
	{
		HRESULT hRes = NULL;
		LPSERVICEADMIN pServiceAdmin = NULL;
		LPSERVICEADMIN2 pServiceAdmin2 = NULL;
		// Retrieves pointers to the supported interfaces on an object.

		CHK_HR_DBG(m_pProfAdmin->AdminServices((LPTSTR)wszProfileName.c_str(), NULL, NULL, MAPI_UNICODE, (LPSERVICEADMIN*)& pServiceAdmin), L"m_pProfAdmin->AdminServices " + wszProfileName);
		CHK_HR_DBG(pServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)& pServiceAdmin2), L"pServiceAdmin->QueryInterface ");

		switch (m_action)
		{
		case ACTION_SERVICE_ADD:
		{
			AddService(pServiceAdmin2);
			break;
		}
		case ACTION_SERVICE_LIST:
		{
			// Get service UID(s) for the services we want to list	
			ULONG cServices = 0;
			if SUCCEEDED(GetABServiceUid(pServiceAdmin2, g_addressBookMap.at(L"displayname").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), g_addressBookMap.at(L"servername").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"servername").c_str(), &cServices, NULL))
			{
				if (cServices > 0)
				{
					MAPIUID* pMAPIUid = new MAPIUID[cServices];
					hRes = GetABServiceUid(pServiceAdmin2, g_addressBookMap.at(L"displayname").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), g_addressBookMap.at(L"servername").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"servername").c_str(), NULL, pMAPIUid), L"Fetching existing service UIDs";
					if ((cServices > 0) && pMAPIUid)
					{
						Logger::WriteLine(LOGLEVEL_INFO, L"Action " + g_toolkitMap.at(L"action") + L" will run against " + ConvertIntToString(cServices) + L" services");
						for (int i = 0; i < cServices; i++)
						{
							if (RunActionOneService(pServiceAdmin2, &pMAPIUid[i]))
								Logger::WriteLine(LOGLEVEL_SUCCESS, L"Address book service succesfully listed");
						}
					}
				}
			}
			break;
		}
		case ACTION_SERVICE_LISTALL:
		{
			ListAllServices(pServiceAdmin2);
			break;
		}

		case ACTION_SERVICE_REMOVEALL:
		{
			RemoveAllServices(pServiceAdmin2);
			break;
		}
		case ACTION_SERVICE_REMOVE:
		{
			switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
			{
			case SERVICETYPE_ADDRESSBOOK:
			{
				// Get service UID(s) for the services we want to remove	
				ULONG cServices = 0;
				if SUCCEEDED(GetABServiceUid(pServiceAdmin2, g_addressBookMap.at(L"displayname").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), g_addressBookMap.at(L"servername").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"servername").c_str(), &cServices, NULL))
				{
					if (cServices > 0)
					{
						MAPIUID* pMAPIUid = new MAPIUID[cServices];
						hRes = GetABServiceUid(pServiceAdmin2, g_addressBookMap.at(L"displayname").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), g_addressBookMap.at(L"servername").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"servername").c_str(), NULL, pMAPIUid), L"Fetching existing service UIDs";
						if ((cServices > 0) && pMAPIUid)
						{
							Logger::WriteLine(LOGLEVEL_INFO, L"Action " + g_toolkitMap.at(L"action") + L" will run against " + ConvertIntToString(cServices) + L" services");
							for (int i = 0; i < cServices; i++)
							{
								if (RunActionOneService(pServiceAdmin2, &pMAPIUid[i]))
									Logger::WriteLine(LOGLEVEL_SUCCESS, L"Address book service succesfully removed");
							}
						}
					}
				}
				break;
			}
			case SERVICETYPE_EXCHANGEACCOUNT:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
			case SERVICETYPE_DATAFILE:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
			}
		}
		case ACTION_SERVICE_UPDATE:
			switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
			{
			case SERVICETYPE_ADDRESSBOOK:
			{
				// Get service UID(s) for the services we want to remove	
				ULONG cServices = 0;
				if SUCCEEDED(GetABServiceUid(pServiceAdmin2, g_addressBookMap.at(L"displayname").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), g_addressBookMap.at(L"servername").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"servername").c_str(), &cServices, NULL))
				{
					if (cServices > 0)
					{
						MAPIUID* pMAPIUid = new MAPIUID[cServices];
						hRes = GetABServiceUid(pServiceAdmin2, g_addressBookMap.at(L"displayname").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"displayname").c_str(), g_addressBookMap.at(L"servername").empty() ? NULL : (LPTSTR)g_addressBookMap.at(L"servername").c_str(), NULL, pMAPIUid), L"Fetching existing service UIDs";
						if ((cServices > 0) && pMAPIUid)
						{
							Logger::WriteLine(LOGLEVEL_INFO, L"Action " + g_toolkitMap.at(L"action") + L" will run against " + ConvertIntToString(cServices) + L" services");
							for (int i = 0; i < cServices; i++)
							{
								if (RunActionOneService(pServiceAdmin2, &pMAPIUid[i]))
									Logger::WriteLine(LOGLEVEL_SUCCESS, L"Address book service succesfully updated");
							}
						}
					}
				}
				break;
			}
			case SERVICETYPE_EXCHANGEACCOUNT:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
			case SERVICETYPE_DATAFILE:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
			}

		}

	Error:
		goto CleanUp;
	CleanUp:
		return SUCCEEDED(hRes);
	}

	BOOL Toolkit::RunActionOneService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID pMapiUid)
	{
		HRESULT hRes = NULL;
		LPPROVIDERADMIN pProviderAdmin = NULL;
		CHK_HR_DBG(pServiceAdmin->AdminProviders(pMapiUid, NULL, &pProviderAdmin), L"Getting profider admin interface pointer for service with UID: " + MapiUidToString(pMapiUid));

		switch (m_action)
		{
			case ACTION_PROVIDER_ADD:
			case ACTION_PROVIDER_UPDATE:
			case ACTION_PROVIDER_LIST:
			case ACTION_PROVIDER_LISTALL:
			case ACTION_PROVIDER_REMOVE:
			case ACTION_PROVIDER_REMOVEALL:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
			case ACTION_SERVICE_UPDATE:
			{
				UpdateService(pServiceAdmin, pMapiUid);
				break;
			}
			case ACTION_SERVICE_SETCACHEDMODE:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
			case ACTION_SERVICE_LIST:
			{
				ListService(pServiceAdmin, pMapiUid);
				break;
			}
			case ACTION_SERVICE_REMOVE:
			{
				RemoveService(pServiceAdmin, pMapiUid);
				break;
			}
			case ACTION_SERVICE_CHANGEDATAFILEPATH:
			case ACTION_SERVICE_SETDEFAULT:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
		}

	Error:
		goto CleanUp;
	CleanUp:
		return SUCCEEDED(hRes);
	}


	VOID Toolkit::ListAllServices(LPSERVICEADMIN2 pServiceAdmin)
	{
		switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
		{
			case SERVICETYPE_ADDRESSBOOK:
			{
				if SUCCEEDED(ListAllABServices(pServiceAdmin))
					Logger::WriteLine(LOGLEVEL_SUCCESS, L"Address book services succesfully listed");
				break;
			}
			case SERVICETYPE_EXCHANGEACCOUNT:
			case SERVICETYPE_DATAFILE:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
		}
	}

	VOID Toolkit::RemoveService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID pMapiUid)
	{
		HRESULT hRes = S_OK;
		switch (g_serviceTypeMap.at(g_toolkitMap.at(L"servicetype")))
		{
			case SERVICETYPE_ADDRESSBOOK:
			{
				CHK_HR_DBG(RemoveABService(pServiceAdmin, pMapiUid), L"RemoveABService");
				break;
			}
			case SERVICETYPE_EXCHANGEACCOUNT:
			case SERVICETYPE_DATAFILE:
			{
				Logger::WriteLine(LOGLEVEL_FAILED, L"The selected action is not currently implemented");
				break;
			}
		}
	Error:
		goto CleanUp;
	CleanUp:
		return;
	}

	BOOL Toolkit::ParseParams(int argc, wchar_t* argv[])
	{
		HRESULT hRes = S_OK;


		// check if we're supposed to list the help menu
		for (int i = 1; i < argc; i++)
		{
			switch (argv[i][0])
			{
			case '-':
			case '/':
			case '\\':
				std::wstring wsArg = SubstringFromStart(1, argv[i]);
				std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);


				if (wsArg == L"?")
				{
					return false;
				}
				break;
			}
		}

		// check if we're supposed to read the configuration from the registry
		for (int i = 1; i < argc; i++)
		{
			switch (argv[i][0])
			{
			case '-':
			case '/':
			case '\\':
				std::wstring wsArg = SubstringFromStart(1, argv[i]);
				std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

				if (wsArg == L"registry")
				{
					m_registry = TRUE;
					ReadConfig();
				}
				break;
			}
		}

		// general toolkit
		for (int i = 1; i < argc; i++)
		{
			switch (argv[i][0])
			{
			case '-':
			case '/':
			case '\\':
				std::wstring wsArg = SubstringFromStart(1, argv[i]);
				std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

				try
				{
					if (i + 1 < argc)
					{
						if ((argv[i + 1][0] != '-') && (argv[i + 1][0] != '/') && (argv[i + 1][0] != '\\'))
						{
							g_toolkitMap.at(wsArg) = argv[i + 1];
							i++;
						}
					};
				}
				catch (const std::exception& e)
				{

				}
				break;
			}
		}

		if (!g_toolkitMap.at(L"profilename").empty())
			g_toolkitMap.at(L"profilemode") = L"specific";

		// If a specific profile is needed then make sure a profile name was specified
		if (VCHK(g_profileModeMap.at(g_toolkitMap.at(L"profilemode")), PROFILEMODE_SPECIFIC) && g_toolkitMap.at(L"profilename").empty())
		{
			Logger::WriteLine(LOGLEVEL_FAILED, L"You must either specify a profile name or pass 'default' for the value of thethe 'profilemode' parameter.");
			return false;
		}

		// address book
		for (int i = 1; i < argc; i++)
		{
			switch (argv[i][0])
			{
			case '-':
			case '/':
			case '\\':
				std::wstring wsArg = SubstringFromStart(1, argv[i]);
				std::transform(wsArg.begin(), wsArg.end(), wsArg.begin(), ::tolower);

				try
				{
					if (i + 1 < argc)
					{
						if ((argv[i + 1][0] != '-') && (argv[i + 1][0] != '/') && (argv[i + 1][0] != '\\'))
						{
							g_addressBookMap.at(wsArg) = argv[i + 1];
							i++;
						}
					};
				}
				catch (const std::exception& e)
				{

				}
				break;
			}
		}





	Error:
		goto CleanUp;
	CleanUp:
		return SUCCEEDED(hRes);
	}
}