/*
* © 2015 Microsoft Corporation
*
* written by Andrei Ghita
*
* Microsoft provides programming examples for illustration only, without warranty either expressed or implied.
* This includes, but is not limited to, the implied warranties of merchantability or fitness for a particular purpose.
* This article assumes that you are familiar with the programming language that is being demonstrated and with
* the tools that are used to create and to debug procedures. Microsoft support engineers can help explain the
* functionality of a particular procedure, but they will not modify these examples to provide added functionality
* or construct procedures to meet your specific requirements.
*/
#include <Windows.h>
#include "..//ToolkitTypeDefs.h"
#include <vector>

namespace MAPIToolkit
{
#pragma region GenericProfile
	LPWSTR GetDefaultProfileNameLP();

	// GetDefaultProfileName
	// returns a std::wstring value with the name of the default Outlook profile
	std::wstring GetDefaultProfileName(LPPROFADMIN lpProfAdmin);

	// GetDefaultProfileName
// returns a std::wstring value with the name of the default Outlook profile
	std::wstring GetDefaultProfileName();

	// GetProfileCount
	// returns the number of mapi profiles for the current user
	ULONG GetProfileCount(LPPROFADMIN lpProfAdmin);

	// GetProfileCount
// returns the number of mapi profiles for the current user
	ULONG GetProfileCount();

	HRESULT HrGetProfiles(ULONG ulProfileCount, ProfileInfo* profileInfo);

	HRESULT HrDeleteProfile(LPWSTR lpszProfileName);

	HRESULT HrCreateProfile(LPWSTR lpszProfileName);

	HRESULT HrCreateProfile(LPWSTR lpszProfileName, LPSERVICEADMIN2* lppSvcAdmin2);

	HRESULT HrSetDefaultProfile(LPWSTR lpszProfileName);

	// Outlook 2016
	HRESULT HrCloneProfile(ProfileInfo* profileInfo);

	// Outlook 2013
	HRESULT HrSimpleCloneProfile(ProfileInfo* profileInfo, bool bSetDefaultProfile);

	VOID PrintProfile(ProfileInfo* profileInfo);

	HRESULT HrGetProfile(LPWSTR lpszProfileName, ProfileInfo* profileInfo);

	HRESULT HrListProfiles(ULONG profileMode, std::wstring profileName, std::wstring wszExportPath);

	BOOL GetProfileNames(LPPROFADMIN pProfAdmin, std::vector<std::wstring>* vProfileNames);

#pragma endregion

#pragma region Providers
	// HrDeleteProvider
	// Deletes the provider with the specified UID from the service with the specified UID in a given profile
	HRESULT HrDeleteProvider(LPWSTR lpwszProfileName, LPMAPIUID lpServiceUid, LPMAPIUID lpProviderUid);

#pragma endregion

#pragma region Sections
	// HrGetSections
	// Returns the EMSMDB and StoreProvider sections of a service
	HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection);

	// HrGetSections
	// Returns the EMSMDB and StoreProvider sections of a service
	HRESULT HrGetSections(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection);
#pragma endregion

#pragma region AddressBook
	HRESULT ListABService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID pMAPIUid);
	HRESULT ListAllABServices(LPSERVICEADMIN2 lpSvcAdmin2);
	HRESULT CreateABService(LPSERVICEADMIN2 lpSvcAdmin2);
	HRESULT UpdateABService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID lpMapiUid);
	HRESULT RemoveABService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID lpMapiUid);
	HRESULT GetABServiceUid(LPSERVICEADMIN2 lpSvcAdmin2, ULONG* ulcMapiUid, MAPIUID* pMapiUid);
	HRESULT GetABServiceUid(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, LPTSTR lppszServerName, ULONG * ulcMapiUid, MAPIUID* pMapiUid);
	HRESULT GetABServiceUid(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, ULONG * ulcMapiUid, MAPIUID* pMapiUid);
	HRESULT CheckABServiceExists(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, LPTSTR lppszServerName, BOOL* success);
	HRESULT CheckABServiceExists(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, BOOL* success);

#pragma endregion

	HRESULT CreateService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID lpuidService);

	STDMETHODIMP CopySBinary(
		LPSBinary psbDest,
		const LPSBinary psbSrc,
		LPVOID pParent);

	HRESULT HrSetABSearchOrder(WCHAR* wszABName, ULONG ulnumEntries, int searchIndex, LPTSTR lpszProfileName);
	HRESULT HrGetABSearchOrderRowCount(LPTSTR lpszProfileName, ULONG * ulnumRows);

	HRESULT HrLogon(LPTSTR lpszProfileName, LPMAPISESSION * lppSession);
	HRESULT HrLogoff(LPMAPISESSION lpSession);
}