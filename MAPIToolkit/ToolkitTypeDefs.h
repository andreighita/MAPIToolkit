#pragma once
#include <string>
#include <Windows.h>
#include <MAPIX.h>
#include <MAPIAux.h>
#include <map>

#define WIN32_LEAN_AND_MEAN

#define ACTION_UNSPECIFIED					0x00000000

#define ACTION_PROFILE_ADD					0x00000001
#define ACTION_PROFILE_CLONE				0x00000002
#define ACTION_PROFILE_RENAME				0x00000004
#define ACTION_PROFILE_LIST					0x00000008
#define ACTION_PROFILE_LISTALL				0x00000010
#define ACTION_PROFILE_REMOVE				0x00000020
#define ACTION_PROFILE_REMOVEALL			0x00000040
#define ACTION_PROFILE_SETDEFAULT			0x00000080
#define ACTION_PROFILE_PROMOTEDELEGATES		0x00000100
#define ACTION_PROFILE_PROMOTEONEDELEGATE	0x00000200

#define ACTION_PROVIDER_ADD					0x00000400
#define ACTION_PROVIDER_UPDATE				0x00000800
#define ACTION_PROVIDER_LIST				0x00001000
#define ACTION_PROVIDER_LISTALL				0x00002000
#define ACTION_PROVIDER_REMOVE				0x00004000
#define ACTION_PROVIDER_REMOVEALL			0x00008000

#define ACTION_SERVICE_ADD					0x00010000
#define ACTION_SERVICE_UPDATE				0x00020000
#define ACTION_SERVICE_SETCACHEDMODE		0x00040000
#define ACTION_SERVICE_LIST					0x00080000 
#define ACTION_SERVICE_LISTALL				0x00100000 
#define ACTION_SERVICE_REMOVE				0x00200000 
#define ACTION_SERVICE_REMOVEALL			0x00400000 
#define ACTION_SERVICE_CHANGEDATAFILEPATH	0x00800000
#define ACTION_SERVICE_SETDEFAULT			0x01000000

//#define ACTION_1								0x02000000 
//#define ACTION_2								0x04000000
//#define ACTION_3								0x08000000
//#define ACTION_4								0x10000000
//#define ACTION_5								0x20000000
//#define ACTION_6								0x40000000
//#define ACTION_7								0x80000000

// PROFILEMODE
#define PROFILEMODE_UNKNOWN		0x00000000
#define PROFILEMODE_DEFAULT		0x00000001
#define PROFILEMODE_SPECIFIC	0x00000002
#define PROFILEMODE_ALL			0x00000004

// SERVICEMODE
#define SERVICEMODE_UNKNOWN		0x00000000
#define SERVICEMODE_DEFAULT		0x00000001
#define SERVICEMODE_SPECIFIC	0x00000002
#define SERVICEMODE_ALL			0x00000004

// CONNECTMODE
#define CONNECTMODE_UNKNOWN			0x00000000
#define CONNECTMODE_RPCOVERHTTP		0x00000001
#define CONNECTMODE_MAPIOVERHTTP	0x00000002

// SERVICETYPE
#define SERVICETYPE_UNKNOWN			0x00000000
#define SERVICETYPE_EXCHANGEACCOUNT	0x00000001
#define SERVICETYPE_DATAFILE		0x00000002
#define SERVICETYPE_ADDRESSBOOK		0x00000004
#define SERVICETYPE_ALL				0x00000008

// PROVIDERTYPE
#define PROVIDERTYPE_UNKNOWN			0x00000000
#define PROVIDERTYPE_PRIMARYMAILBOX		0x00000001
#define PROVIDERTYPE_DELEGATE			0x00000002
#define PROVIDERTYPE_PUBLICFOLDERS		0x00000004
#define PROVIDERTYPE_ALL				0x00000008

// EXPORTMODE
#define EXPORTMODE_UNKNOWN			0x00000000
#define EXPORTMODE_EXPORT			0x00000001
#define EXPORTMODE_NOEXPORT			0x00000002

// INPUTMODE
#define INPUTMODE_UNKNOWN			0x00000000
#define INPUTMODE_USER				0x00000001
#define INPUTMODE_DIRECTORY			0x00000002
#define INPUTMODE_FILE				0x00000004

// LOGGINGMODE
#define LOGGINGMODE_UNKNOWN			0x00000000
#define LOGGINGMODE_NONE			0x00000001
#define LOGGINGMODE_CONSOLE			0x00000002
#define LOGGINGMODE_FILE			0x00000004
#define LOGGINGMODE_ALL				0x00000008
#define LOGGINGMODE_DEBUG			0x00000010

// CACHEDMODE
#define CACHEDMODE_UNKNOWN			0x00000000
#define CACHEDMODE_ENABLED			0x00000001
#define CACHEDMODE_DISABLED			0x00000002

// LOGLEVEL
#define LOGLEVEL_UNKNOWN			0x00000000
#define LOGLEVEL_INFO				0x00000001
#define LOGLEVEL_WARNING			0x00000002
#define LOGLEVEL_ERROR				0x00000004
#define LOGLEVEL_SUCCESS			0x00000008
#define LOGLEVEL_FAILED				0x00000010
#define LOGLEVEL_DEBUG				0x00000020

// LOGCALLSTATUS
#define LOGCALLSTATUS_UNKNOWN			0x00000000
#define LOGCALLSTATUS_SUCCESS			0x00000001
#define LOGCALLSTATUS_ERROR				0x00000002
#define LOGCALLSTATUS_NOFILE			0x00000004
#define LOGCALLSTATUS_LOGGINGDISABLED	0x00000008

// PSTTYPE
#define PSTTYPE_UNKNOWN			0x00000000
#define PSTTYPE_ANSI			0x00000001
#define PSTTYPE_UNICODE			0x00000002

// SAVECONFIG
#define SAVECONFIG_UNKNOWN		0x00000000
#define SAVECONFIG_TRUE			0x00000001
#define SAVECONFIG_FALSE		0x00000002

struct UpdateSmtpAddress
{
	ULONG ulAdTimeout;
	ULONG inputMode;
	std::wstring szADsPAth;
	std::wstring szOldDomainName;
	std::wstring szNewDomainName;
};

struct MailboxInfo
{
	std::wstring wszDisplayName; // PR_DISPLAY_NAME
	std::wstring wszSmtpAddress; // PR_PROFILE_USER_SMTP_EMAIL_ADDRESS
	std::wstring wszProfileMailbox; // PR_PROFILE_MAILBOX
	std::wstring wszProfileServerDN; // PR_PROFILE_SERVER_DN
	std::wstring wszRohProxyServer; // PR_ROH_PROXY_SERVER
	std::wstring wszProfileServer; // PR_PROFILE_SERVER
	std::wstring wszProfileServerFqdnW; // PR_PROFILE_SERVER_FQDN_W
	std::wstring wszAutodiscoverUrl; // PR_PROFILE_LKG_AUTODISCOVER_URL
	std::wstring wszMailStoreInternalUrl; // PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
	std::wstring wszMailStoreExternalUrl; // PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
	std::wstring wszAddressBookInternalUrl; // PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL
	std::wstring wszAddressBookExternalUrl; // PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL
	BOOL bPrimaryMailbox;
	ULONG ulResourceFlags; // PR_RESOURCE_FLAGS
	ULONG ulRohProxyAuthScheme; // PR_PROFILE_RPC_PROXY_SERVER_AUTH_PACKAGE
	ULONG ulRohFlags; // PR_ROH_FLAGS
	ULONG ulProfileType; // PR_PROFILE_TYPE
	MAPIUID muidProviderUid;
	MAPIUID muidServiceUid;
	BOOL bIsOnlineArchive;
};

struct PstInfo
{
	std::wstring wszDisplayName;
	std::wstring wszPstPath; // PR_PST_PATH_W
	ULONG ulPstType;
	ULONG ulPstConfigFlags; // PR_PST_CONFIG_FLAGS
};

struct MapiProperty
{
	std::wstring wszNamedPropertyName;
	std::wstring wszPropertyTag;
	ULONG ulNamedPropertyValue;
};

struct ProviderInfo
{
	MapiProperty* mapiProperties;
};

struct ExchangeAccountInfo
{
	std::wstring wszDisplayName;
	std::wstring wszDatafilePath;
	std::wstring wszEmailAddress;
	BOOL bULONGEnabledOwner;
	BOOL bULONGEnabledShared;
	BOOL bULONGEnabledPublicFolders;
	int iCachedModeMonths;
	std::wstring szUserName;
	std::wstring szUserEmailSmtpAddress;
	ULONG ulMailboxCount;
	std::wstring wszRohProxyServer;
	std::wstring wszUnresolvedServer;
	std::wstring wszHomeServerName;
	std::wstring wszHomeServerDN;
	std::wstring wszMailboxDN;
	MailboxInfo* accountMailboxes;
	ULONG ulProfileConfigFlags;
	std::wstring wszMailStoreInternalUrl; // PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
	std::wstring wszMailStoreExternalUrl; // PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL
	std::wstring wszAddressBookInternalUrl; // PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL
	std::wstring wszAddressBookExternalUrl; // PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL

};

struct EMSMdbSection
{
	std::wstring wszDisplayName;
	std::wstring wszDatafilePath;
	std::wstring wszSmtpAddress;
	BOOL bULONGEnabledOwner;
	BOOL bULONGEnabledShared;
	BOOL bULONGEnabledPublicFolders;
	int iCachedModeMonths;
	std::wstring szUserName;
	std::wstring szUserEmailSmtpAddress;
	ULONG ulMailboxCount;
	std::wstring wszRohProxyServer;
	std::wstring wszUnresolvedServer;
	std::wstring wszHomeServerName;
	std::wstring wszHomeServerDN;
	MailboxInfo* accountMailboxes;
	ULONG ulProfileConfigFlags;
};

struct AddressBookProviderInfo
{
	std::wstring szDisplayName;
	std::wstring szAbServerName;
	std::wstring szAbUsername;
	ULONG ulMaxEntries;
	ULONG ulTimeout;
	ULONG ulSlowTimeout;
	BOOL bUseSSL;
	BOOL bUsePSA;
	ULONG ulAbServerPort;
};

struct EasAccountInfo
{
	std::wstring szDisplayName;
	std::wstring szDataFilePath; // PR_PROFILE_OFFLINE_STORE_PATH_W
};

struct OtherServiceInfo
{
	std::wstring szDisplayName;
	std::wstring szServiceName;
};

struct ServiceInfo
{
	std::wstring wszServiceName;
	std::wstring wszDisplayName;
	ULONG serviceType; // MSEMS = SERVICETYPE_EXCHANGEACCOUNT; EMABLT = SERVICETYPE_ADDRESSBOOKPROVIDER; MSPST_MS/MSUPST_MS = ULONG::Pst; EAS = SERVICETYPE_EASACCOUNT;
	BOOL bDefaultStore; // PR_RESOURCE_FLAGS = 
	ULONG ulResourceFlags; // PR_RESOURCE_FLAGS = SERVICE_DEFAULT_STORE
	EasAccountInfo* easAccountInfo;
	ExchangeAccountInfo* exchangeAccountInfo;
	AddressBookProviderInfo* addressBookProviderInfo;
	PstInfo* pstInfo;
	MAPIUID muidServiceUid;
};

struct ProfileInfo
{
	std::wstring wszProfileName = L"";
	BOOL bDefaultProfile = false;
	ULONG ulServiceCount = 0;
	ServiceInfo* profileServices = NULL;
};

#define AB_PROVIDER_BASE_ID						0x6600  // Look at the comments in MAPITAGS.H
#define PROP_AB_PROVIDER_DISPLAY_NAME			PR_DISPLAY_NAME_A
#define PROP_AB_PROVIDER_SERVER_NAME			PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0000))	// "example.contoso.com"
#define PROP_AB_PROVIDER_SERVER_PORT			PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0001)) // "389"
#define PROP_AB_PROVIDER_USER_NAME				PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0002)) // contoso\administrator
#define PROP_AB_PROVIDER_SEARCH_BASE			PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0003)) // SEARCH_FILTER_VALUE
#define PROP_AB_PROVIDER_SEARCH_FILTER			PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0004)) // "(&(mail=*)(|(mail=%s*)(|(cn=%s*)(|(sn=%s*)(givenName=%s*)))))"
#define PROP_AB_PROVIDER_ADDRTYPE				PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0005))	// "SMTP"
#define PROP_AB_PROVIDER_SOURCE					PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0006))	// "mail"
#define PROP_AB_PROVIDER_SEARCH_TIMEOUT			PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0007))	// "60"
#define PROP_AB_PROVIDER_MAX_ENTRIES			PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0008)) // "100"
#define PROP_AB_PROVIDER_SEARCH_TIMEOUT_LBW		PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0009))	// "120"
#define PROP_AB_PROVIDER_MAX_ENTRIES_LBW		PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x000a)) // "15"
#define PROP_AB_PROVIDER_LOGFILE				PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x000b)) // ""
#define PROP_AB_PROVIDER_ERRLOGGING				PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x000c)) // "OFF"
#define PROP_AB_PROVIDER_DIAGTRACING			PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x000d)) // "OFF"
#define PROP_AB_PROVIDER_TRACELEVEL				PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x000e)) // "NONE"
#define PROP_AB_PROVIDER_DEBUGWIN				PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x000f)) // "OFF"
#define PROP_AB_PROVIDER_ADDITIONAL_INFO_SOURCE	PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0010)) // "postalAddress"
#define PROP_AB_PROVIDER_DISPLAY_NAME_SOURCE	PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0011)) // "cn"
#define PROP_AB_PROVIDER_LDAP_UI				PROP_TAG (PT_STRING8,	(AB_PROVIDER_BASE_ID + 0x0012)) // "1"
#define PROP_AB_PROVIDER_USE_SSL				PROP_TAG (PT_BOOLEAN,	(AB_PROVIDER_BASE_ID + 0x0013)) // False
#define PROP_AB_PROVIDER_SERVER_SPA				PROP_TAG (PT_BOOLEAN,	(AB_PROVIDER_BASE_ID + 0x0015)) // False
#define PROP_AB_PROVIDER_USER_PASSWORD_ENCODED	PROP_TAG (PT_BINARY,	(AB_PROVIDER_BASE_ID + 0x0017)) // ENCODED_PWD
#define PROP_AB_PROVIDER_ENABLE_BROWSING		PROP_TAG(PT_BOOLEAN,	(AB_PROVIDER_BASE_ID + 0x0022)) // False
#define PROP_AB_PROVIDER_SEARCH_BASE_DEFAULT	PROP_TAG(PT_LONG,		(AB_PROVIDER_BASE_ID + 0x0023)) // 0 or 1

struct ABProvider
{
	LPTSTR lpszDisplayName;  // LPTSTR = LPWSTR; LPSTR  
	LPTSTR lpszServerName;
	LPTSTR lpszServerPort;
	BOOL bUseSSL;
	LPTSTR lpszUsername;
	LPTSTR lpszPassword;
	BOOL bRequireSPA;
	LPTSTR lpszTimeout;
	LPTSTR lpszMaxResults;
	ULONG ulDefaultSearchBase;
	LPTSTR lpszCustomSearchBase;
	BOOL bEnableBrowsing;
	LPTSTR lpszServiceName;
};

// PST
#define PR_PST_CONFIG_FLAGS PROP_TAG(PT_LONG, 0x6770)
#define PR_PST_PATH_W PROP_TAG(PT_UNICODE, 0x6700)