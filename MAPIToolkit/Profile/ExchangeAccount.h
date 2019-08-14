#pragma once
#include "../ExchangeAccountWorker.h"

namespace MAPIToolkit
{
	HRESULT HrCreateMsemsService(ULONG profileMode, LPWSTR lpwszProfileName, int iOutlookVersion, ExchangeAccountWorker* pExchangeAccountWorker);

	HRESULT HrCreateMsemsServiceOneProfile(LPWSTR lpwszProfileName, int iOutlookVersion, ExchangeAccountWorker* pExchangeAccountWorker);

	// HrCrateMsemsServiceModernExt
	// Crates a new message store service and configures the following properties:
	// - PR_PROFILE_CONFIG_FLAGS
	// - PR_RULE_ACTION_TYPE
	// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
	// - PR_DISPLAY_NAME_W
	// - PR_PROFILE_ACCT_NAME_W
	// - PR_PROFILE_UNRESOLVED_NAME_W
	// - PR_PROFILE_USER_EMAIL_W
	// Also updates the store provider section with the two following properties:
	// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
	// - PR_DISPLAY_NAME_W
	// This implementation is Outlook 2016 specific
	HRESULT HrCreateMsemsServiceModernExt(BOOL bDefaultProfile,
		LPWSTR lpwszProfileName,
		ULONG ulResourceFlags,
		ULONG ulProfileConfigFlags,
		ULONG ulULONGMonths,
		LPWSTR lpszSmtpAddress,
		LPWSTR lpszDisplayName);

	// HrCrateMsemsServiceModern
	// Crates a new message store service and configures the following properties:
	// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
	// - PR_DISPLAY_NAME_W
	// - PR_PROFILE_ACCT_NAME_W
	// - PR_PROFILE_UNRESOLVED_NAME_W
	// - PR_PROFILE_USER_EMAIL_W
	// Also updates the store provider section with the two following properties:
	// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
	// - PR_DISPLAY_NAME_W
	// This implementation is Outlook 2016 specific
	HRESULT HrCreateMsemsServiceModern(BOOL bDefaultProfile,
		LPWSTR lpwszProfileName,
		LPWSTR lpszSmtpAddress,
		LPWSTR lpszDisplayName);

	// HrCreateMsemsServiceLegacyUnresolved
	// Crates a new message store service and configures the following properties it with a default property set. 
	// This is the legacy implementation where Outlook resolves the mailbox based on "unresolved" mailbox and server names. I use this for Outlook 2007.
	HRESULT HrCreateMsemsServiceLegacyUnresolved(BOOL bDefaultProfile,
		LPWSTR lpwszProfileName,
		LPWSTR lpszwMailboxDN,
		LPWSTR lpszwServer);

	// HrCreateMsemsServiceROH
	//Creates a new message store service and sets it for RPC / HTTP with the following properties:
	//	PR_PROFILE_USER
	//	PR_DISPLAY_NAME_W
	//	PR_PROFILE_UNRESOLVED_NAME_W
	//	PR_PROFILE_HOME_SERVER
	//	PR_PROFILE_HOME_SERVER_FQDN
	//	PR_PROFILE_HOME_SERVER_DN
	//	PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
	//	PR_PROFILE_HOME_SERVER_ADDRS
	//	PR_PROFILE_ACCT_NAME_W
	//	PR_PROFILE_CONFIG_FLAGS
	//	PR_PROFILE_TRANSPORT_FLAGS
	//	PR_PROFILE_CONNECT_FLAGS
	//	PR_PROFILE_UI_STATE
	//	PR_PROFILE_AUTH_PACKAGE
	//Configures the Store Provider with the following properties:
	//	PR_PROFILE_SERVER
	//	PR_PROFILE_SERVER_FQDN
	//	PR_PROFILE_SERVER_DN
	//	PR_PROFILE_MAILBOX
	//	PR_DISPLAY_NAME_W
	//	PR_PROFILE_DISPLAYNAME_SET
	//	PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W 
	HRESULT HrCreateMsemsServiceROH(BOOL bDefaultProfile,
		LPWSTR lpwszProfileName,
		LPWSTR lpszSmtpAddress,
		LPWSTR lpszMailboxLegacyDn,
		LPWSTR lpszUnresolvedServer,
		LPWSTR lpszRohProxyServer,
		LPWSTR lpszProfileServerDn,
		LPWSTR lpszAutodiscoverUrl);

	// HrCreateMsemsServiceMOH
	// Creates a new message store service and sets it for MAPI / HTTP with the following properties:
	// - PR_PROFILE_CONFIG_FLAGS
	// - PR_PROFILE_AUTH_PACKAGE
	// - PR_PROFILE_MAPIHTTP_ADDRESSBOOK_INTERNAL_URL
	// - PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL
	// - PR_PROFILE_USER
	// Configures the Store Provider with the following properties:
	// - PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W
	// - PR_DIPLAY_NAME_W
	// - PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL
	// - PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
	HRESULT HrCreateMsemsServiceMOH(BOOL bDefaultProfile,
		LPWSTR lpwszProfileName,
		LPWSTR lpszSmtpAddress,
		LPWSTR lpszMailboxDn,
		LPWSTR lpszServerDn,
		LPWSTR lpszServerName,
		LPWSTR lpszMailStoreInternalUrl,
		LPWSTR lpszMailStoreExternalUrl,
		LPWSTR lpszAddressBookInternalUrl,
		LPWSTR lpszAddressBookExternalUrl);

	// HrGetDefaultMsemsServiceAdminProviderPtr
	// Returns the provider admin interface pointer for the default service in a given profile
	HRESULT HrGetDefaultMsemsServiceAdminProviderPtr(LPWSTR lpwszProfileName, LPPROVIDERADMIN* lppProvAdmin, LPMAPIUID* lppServiceUid);

	// HrGetSections
	// Returns the EMSMDB and StoreProvider sections of a service
	HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection);

	// HrGetSections
	// Returns the EMSMDB and StoreProvider sections of a service
	HRESULT HrGetSections(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection);

	HRESULT HrUpdatePrStoreProviders(LPSERVICEADMIN lpServiceAdmin, LPMAPIUID lpServiceUid, LPMAPIUID lpProviderUid);

}