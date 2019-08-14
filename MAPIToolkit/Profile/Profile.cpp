#include "Profile.h"
#include "..//Logger.h"
#include "../InlineAndMacros.h"
#include <MAPIUtil.h>
#include <EdkMdb.h>
#include "../ExtraMAPIDefs.h"
#include "..//Misc/Utility/StringOperations.h"
#include <string.h>
#include <wincrypt.h>
#include "..//Misc/XML/XMLHelper.h"
#include "AdditionalMailbox.h"
#include "ExchangeAccount.h"
#include "PST.h"
#include "..//InlineAndMacros.h"
#include "..//Toolkit.h"
#include <wchar.h> 
namespace MAPIToolkit
{
LPWSTR GetDefaultProfileNameLP()
{
	return (LPWSTR)GetDefaultProfileName().c_str();
}

HRESULT HrListProfiles(ULONG profileMode, std::wstring profileName, std::wstring wszExportPath)
{
	HRESULT hRes = S_OK;
	if VCHK(profileMode, PROFILEMODE_ALL)
	{
		ULONG ulProfileCount = GetProfileCount();
		ProfileInfo* profileInfo = new ProfileInfo[ulProfileCount];
		ZeroMemory(profileInfo, sizeof(ProfileInfo) * ulProfileCount);
		Logger::Write(LOGLEVEL_INFO, L"Retrieving MAPI Profile information for all profiles");
		CHK_HR_DBG(HrGetProfiles(ulProfileCount, profileInfo), L"Calling HrGetProfiles");
		if (wszExportPath != L"")
		{
			Logger::Write(LOGLEVEL_INFO, L"Exporting MAPI Profile information for all profiles");
			ExportXML(ulProfileCount, profileInfo, wszExportPath);
		}
		else
		{
			Logger::Write(LOGLEVEL_INFO, L"Exporting MAPI Profile information for all profiles");
			ExportXML(ulProfileCount, profileInfo, L"");
		}

	}
	else if VCHK(profileMode, PROFILEMODE_SPECIFIC)
	{
		ProfileInfo* pProfileInfo = new ProfileInfo();
		Logger::Write(LOGLEVEL_INFO, L"Retrieving MAPI Profile information for profile: " + profileName);
		CHK_HR_DBG(HrGetProfile((LPWSTR)profileName.c_str(), pProfileInfo), L"Calling HrGetProfile");
		if (wszExportPath != L"")
		{
			Logger::Write(LOGLEVEL_INFO, L"Exporting MAPI Profile information for profile");
			ExportXML(1, pProfileInfo, wszExportPath);
		}
		else
		{
			Logger::Write(LOGLEVEL_INFO, L"Exporting MAPI Profile information for profile");
			ExportXML(1, pProfileInfo, L"");
		}

	}
	else if VCHK(profileMode, PROFILEMODE_DEFAULT)
	{
		std::wstring szDefaultProfileName = GetDefaultProfileName();
		if (!szDefaultProfileName.empty())
		{
			profileName = szDefaultProfileName;
		}

		ProfileInfo* pProfileInfo = new ProfileInfo();
		Logger::Write(LOGLEVEL_INFO, L"Retrieving MAPI Profile information for default profile: " + profileName);
		CHK_HR_DBG(HrGetProfile((LPWSTR)profileName.c_str(), pProfileInfo), L"Calling HrGetProfile");
		if (wszExportPath != L"")
		{
			Logger::Write(LOGLEVEL_INFO, L"Exporting MAPI Profile information for default profile");
			ExportXML(1, pProfileInfo, wszExportPath);
		}
		else
		{
			Logger::Write(LOGLEVEL_INFO, L"Exporting MAPI Profile information for default profile");
			ExportXML(1, pProfileInfo, L"");
		}
	}
Error:
	goto CleanUp;
CleanUp:
	return hRes;
}

BOOL GetProfileNames(LPPROFADMIN pProfAdmin, std::vector<std::wstring>* vProfileNames)
{
	HRESULT hRes = S_OK;
	std::vector<std::wstring> vProfileNamesTemp;
	LPMAPITABLE pMapiTable = NULL;
	LPSRowSet pSRowSet = NULL;

	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A };

	CHK_HR_DBG(pProfAdmin->GetProfileTable(0, &pMapiTable), L"Getting the profile table");
	

	// Query the table to get the the default profile only
	CHK_HR_DBG(HrQueryAllRows(pMapiTable,
		(LPSPropTagArray)& sptaProps,
		NULL,
		NULL,
		0,
		&pSRowSet), L"Calling HrQueryAllRows");

	if (pSRowSet->cRows == 0)
	{
		Logger::Write(LOGLEVEL_FAILED, L"No profiles found.");
	}
	else 
	{
		for (int i = 0; i < pSRowSet->cRows; i++)
		{
			vProfileNamesTemp.push_back(ConvertMultiByteToStdWString(pSRowSet->aRow[i].lpProps[iDisplayName].Value.lpszA));
		}
	}

Error:
	goto CleanUp;
CleanUp:
	if (pMapiTable) pMapiTable->Release();
	if (pSRowSet) FreeProws(pSRowSet);
	vProfileNames = &vProfileNamesTemp;
	return SUCCEEDED(hRes);
}

// GetDefaultProfileName
// returns a std::wstring value with the name of the default Outlook profile
std::wstring GetDefaultProfileName(LPPROFADMIN lpProfAdmin)
{
	std::wstring szDefaultProfileName;
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	HRESULT hRes = S_OK;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A };

	CHK_HR_DBG(lpProfAdmin->GetProfileTable(0, &lpProfTable), L"Calling GetProfileTable");

	// Allocate memory for the restriction
	CHK_HR_DBG(MAPIAllocateBuffer(sizeof(SRestriction), (LPVOID*)&lpProfRes), L"Calling MAPIAllocateBuffer");

	CHK_HR_DBG(MAPIAllocateBuffer(sizeof(SRestriction) * 2, (LPVOID*)&lpProfResLvl1), L"Calling MAPIAllocateBuffer");

	CHK_HR_DBG(MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)&lpProfPropVal), L"Calling MAPIAllocateBuffer");

	// Set up restriction to query the profile table
	lpProfRes->rt = RES_AND;
	lpProfRes->res.resAnd.cRes = 0x00000002;
	lpProfRes->res.resAnd.lpRes = lpProfResLvl1;

	lpProfResLvl1[0].rt = RES_EXIST;
	lpProfResLvl1[0].res.resExist.ulPropTag = PR_DEFAULT_PROFILE;
	lpProfResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
	lpProfResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
	lpProfResLvl1[1].rt = RES_PROPERTY;
	lpProfResLvl1[1].res.resProperty.relop = RELOP_EQ;
	lpProfResLvl1[1].res.resProperty.ulPropTag = PR_DEFAULT_PROFILE;
	lpProfResLvl1[1].res.resProperty.lpProp = lpProfPropVal;

	lpProfPropVal->ulPropTag = PR_DEFAULT_PROFILE;
	lpProfPropVal->Value.b = true;

	// Query the table to get the the default profile only
	CHK_HR_DBG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), L"Calling HrQueryAllRows");

	if (lpProfRows->cRows == 0)
	{
		Logger::Write(LOGLEVEL_FAILED, L"No default profile set.");
	}
	else if (lpProfRows->cRows == 1)
	{

		szDefaultProfileName = ConvertMultiByteToWideChar(lpProfRows->aRow->lpProps[iDisplayName].Value.lpszA);
	}
	else
	{
		Logger::Write(LOGLEVEL_ERROR, L"Query resulted in incosinstent results");
	}

Error:
	goto CleanUp;
CleanUp:
	// Free up memory
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfRes) MAPIFreeBuffer(lpProfRes);
	if (lpProfResLvl1) MAPIFreeBuffer(lpProfResLvl1);
	return szDefaultProfileName;
}

// GetDefaultProfileName
// returns a std::wstring value with the name of the default Outlook profile
std::wstring GetDefaultProfileName()
{
	std::wstring szDefaultProfileName;
	LPPROFADMIN lpProfAdmin = NULL;
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	HRESULT hRes = S_OK;


	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A };
	CHK_HR_DBG(MAPIAdminProfiles(0, &lpProfAdmin), L"MAPIAdminProfiles");

	CHK_HR_DBG(lpProfAdmin->GetProfileTable(0, &lpProfTable), L"Calling GetProfileTable");

	// Allocate memory for the restriction
	CHK_HR_DBG(MAPIAllocateBuffer(sizeof(SRestriction), (LPVOID*)& lpProfRes), L"Calling MAPIAllocateBuffer");

	CHK_HR_DBG(MAPIAllocateBuffer(sizeof(SRestriction) * 2, (LPVOID*)& lpProfResLvl1), L"Calling MAPIAllocateBuffer");

	CHK_HR_DBG(MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)& lpProfPropVal), L"Calling MAPIAllocateBuffer");

	// Set up restriction to query the profile table
	lpProfRes->rt = RES_AND;
	lpProfRes->res.resAnd.cRes = 0x00000002;
	lpProfRes->res.resAnd.lpRes = lpProfResLvl1;

	lpProfResLvl1[0].rt = RES_EXIST;
	lpProfResLvl1[0].res.resExist.ulPropTag = PR_DEFAULT_PROFILE;
	lpProfResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
	lpProfResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
	lpProfResLvl1[1].rt = RES_PROPERTY;
	lpProfResLvl1[1].res.resProperty.relop = RELOP_EQ;
	lpProfResLvl1[1].res.resProperty.ulPropTag = PR_DEFAULT_PROFILE;
	lpProfResLvl1[1].res.resProperty.lpProp = lpProfPropVal;

	lpProfPropVal->ulPropTag = PR_DEFAULT_PROFILE;
	lpProfPropVal->Value.b = true;

	// Query the table to get the the default profile only
	CHK_HR_DBG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)& sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), L"Calling HrQueryAllRows");

	if (lpProfRows->cRows == 0)
	{
		Logger::Write(LOGLEVEL_FAILED, L"No default profile set.");
	}
	else if (lpProfRows->cRows == 1)
	{

		szDefaultProfileName = ConvertMultiByteToWideChar(lpProfRows->aRow->lpProps[iDisplayName].Value.lpszA);
	}
	else
	{
		Logger::Write(LOGLEVEL_ERROR, L"Query resulted in incosinstent results");
	}

Error:
	goto CleanUp;
CleanUp:
	// Free up memory
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfRes) MAPIFreeBuffer(lpProfRes);
	if (lpProfResLvl1) MAPIFreeBuffer(lpProfResLvl1);
	if (lpProfAdmin) lpProfAdmin->Release();
	return szDefaultProfileName;
}


// GetProfileCount
// returns the number of mapi profiles for the current user
ULONG GetProfileCount()
{
	std::string szDefaultProfileName;
	LPMAPITABLE lpProfTable = NULL;
	ULONG ulRowCount = 0;
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;
	CHK_HR_DBG(MAPIAdminProfiles(0, &lpProfAdmin), L"MAPIAdminProfiles");
	CHK_HR_DBG(lpProfAdmin->GetProfileTable(0, &lpProfTable), L"lpProfAdmin->GetProfileTable");
	CHK_HR_DBG(lpProfTable->GetRowCount(0, &ulRowCount), L"lpProfTable->GetRowCount");

Error:
	goto CleanUp;
CleanUp:
// Free up memory
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return ulRowCount;
}

ULONG GetProfileCount(LPPROFADMIN lpProfAdmin)
{
	std::string szDefaultProfileName;
	LPMAPITABLE lpProfTable = NULL;
	ULONG ulRowCount = 0;
	HRESULT hRes = S_OK;
	CHK_HR_DBG(lpProfAdmin->GetProfileTable(0, &lpProfTable), L"lpProfAdmin->GetProfileTable");
	CHK_HR_DBG(lpProfTable->GetRowCount(0, &ulRowCount), L"lpProfTable->GetRowCount");

Error:
	goto CleanUp;
CleanUp:
	// Free up memory
	if (lpProfTable) lpProfTable->Release();
	return ulRowCount;
}

HRESULT HrGetProfiles(ULONG ulProfileCount, ProfileInfo * profileInfo)
{
	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;

	HRESULT hRes = S_OK;

	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME_A };

	CHK_HR_DBG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
									 // Get an IProfAdmin interface.

	CHK_HR_DBG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), L"Calling GetProfileTable");

	// Query the table to get the the default profile only
	CHK_HR_DBG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		NULL,
		NULL,
		0,
		&lpProfRows), L"Calling HrQueryAllRows");

	if (lpProfRows->cRows == ulProfileCount)
	{
		for (unsigned int i = 0; i < lpProfRows->cRows; i++)
		{
			CHK_HR_DBG(HrGetProfile(ConvertMultiByteToWideChar(lpProfRows->aRow[i].lpProps[iDisplayName].Value.lpszA), &profileInfo[i]), L"Calling HrGetProfile");
		}
	}

Error:
	goto CleanUp;
CleanUp:
	// Free up memory
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();
	return hRes;
}

HRESULT HrDeleteProfile(LPWSTR lpszProfileName)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.
	LPSERVICEADMIN		lpSvcAdmin = NULL;      // Service Admin object.
	LPSERVICEADMIN2		lpSvcAdmin2 = NULL;

	// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	CHK_HR_DBG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
		CHK_HR_DBG(lpProfAdmin->DeleteProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName), NULL), L"Calling DeleteProfile");
		// Create a new profile.

	// Clean up
	if (lpProfAdmin) lpProfAdmin->Release();

Error:
	goto CleanUp;
CleanUp:
	return 0;

}

HRESULT HrCreateProfile(LPWSTR lpszProfileName)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.
	LPSERVICEADMIN		lpSvcAdmin = NULL;      // Service Admin object.
	LPSERVICEADMIN2		lpSvcAdmin2 = NULL;

	// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	CHK_HR_DBG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
	hRes = lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
		nullptr,          // Password for profile.
		0,          // Handle to parent window.
		0);        // Flags.

	if (hRes == E_ACCESSDENIED)
	{
		CHK_HR_DBG(lpProfAdmin->DeleteProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName), NULL), L"Calling DeleteProfile");
		// Create a new profile.

		CHK_HR_DBG(lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
			nullptr,          // Password for profile.
			0,          // Handle to parent window.
			0), L"Calling CreateProfile.");        // Flags.
	}

	// Clean up
	if (lpProfAdmin) lpProfAdmin->Release();

Error:
	goto CleanUp;
CleanUp:
	return 0;

}

HRESULT HrCreateProfile(LPWSTR lpszProfileName, LPSERVICEADMIN2 *lppSvcAdmin2)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.
	LPSERVICEADMIN		lpSvcAdmin = NULL;      // Service Admin object.
	LPSERVICEADMIN2		lpSvcAdmin2 = NULL;

	// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	CHK_HR_DBG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
	hRes = lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
		nullptr,          // Password for profile.
		0,          // Handle to parent window.
		0);        // Flags.

	if (hRes == E_ACCESSDENIED)
	{
		CHK_HR_DBG(lpProfAdmin->DeleteProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName), NULL), L"Calling DeleteProfile.");

		CHK_HR_DBG(lpProfAdmin->CreateProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
			nullptr,          // Password for profile.
			0,          // Handle to parent window.
			0), L"Calling CreateProfile.");        // Flags.
	}

	// Get an IMsgServiceAdmin interface off of the IProfAdmin interface.
	CHK_HR_DBG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Profile that we want to modify.
		nullptr,          // Password for that profile.
		0,          // Handle to parent window.
		0,             // Flags.
		&lpSvcAdmin), L"Calling AdminServices."); // Pointer to new IMsgServiceAdmin.

												  // Create the new message service for Exchange.
	if (lpSvcAdmin) CHK_HR_DBG(lpSvcAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpSvcAdmin2), L"Calling QueryInterface");

	*lppSvcAdmin2 = lpSvcAdmin2;

Error:
	goto CleanUp;

CleanUp:
	// Clean up
	if (lpSvcAdmin) lpSvcAdmin->Release();
	if (lpProfAdmin) lpProfAdmin->Release();

	return 0;

}

HRESULT HrSetDefaultProfile(LPWSTR lpszProfileName)
{
	HRESULT				hRes = S_OK;            // Result from MAPI calls.
	LPPROFADMIN			lpProfAdmin = NULL;     // Profile Admin object.

												// This indicates columns we want returned from HrQueryAllRows.
	enum { iSvcName, iSvcUID, cptaSvc };
	SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

	// Get an IProfAdmin interface.

	CHK_HR_DBG(MAPIAdminProfiles(0,              // Flags.
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin.

													   // Create a new profile.
	CHK_HR_DBG(lpProfAdmin->SetDefaultProfile((LPTSTR)ConvertWideCharToMultiByte(lpszProfileName),     // Name of new profile.
		0), L"Calling SetDefaultProfile.");        // Flags.

	// Clean up
	if (lpProfAdmin) lpProfAdmin->Release();
Error:
	goto CleanUp;
CleanUp:
	return 0;

}

// Outlook 2016
HRESULT HrCloneProfile(ProfileInfo * profileInfo)
{
	HRESULT hRes = S_OK;
	LPSERVICEADMIN2 lpServiceAdmin = NULL;
	unsigned int uiServiceIndex = 0;
	profileInfo->wszProfileName = profileInfo->wszProfileName + L"_Clone";
	Logger::Write(LOGLEVEL_INFO, L"Creating new profile named: " + profileInfo->wszProfileName);
	CHK_HR_DBG(HrCreateProfile((LPWSTR)profileInfo->wszProfileName.c_str(), &lpServiceAdmin), L"Calling HrCreateProfile.");
	if (lpServiceAdmin)
	{
		for (unsigned int i = 0; i < profileInfo->ulServiceCount; i++)
		{
			MAPIUID uidService = { 0 };
			LPMAPIUID lpServiceUid = &uidService;
			if (profileInfo->profileServices[i].serviceType == SERVICETYPE_EXCHANGEACCOUNT)
			{
				Logger::Write(LOGLEVEL_INFO, L"Adding exchange mailbox: " + profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress);
				CHK_HR_DBG(HrCreateMsemsServiceModernExt(false, // sort this out later
					(LPWSTR)profileInfo->wszProfileName.c_str(),
					profileInfo->profileServices[i].ulResourceFlags,
					profileInfo->profileServices[i].exchangeAccountInfo->ulProfileConfigFlags,
					profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths,
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName.c_str()), L"Calling HrCreateMsemsServiceModernExt.");

				MAPIUID uidService = { 0 };
				memcpy((LPVOID)&uidService, lpServiceUid, sizeof(MAPIUID));
				for (unsigned int j = 0; j < profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
				{
					if (profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType == PROFILE_DELEGATE)
					{
						Logger::Write(LOGLEVEL_INFO, L"Adding additional mailbox: " + profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress);
						// this should not add online archives
						if (TRUE != profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive)
							CHK_HR_DBG(HrAddDelegateMailboxModern(false,
							(LPWSTR)profileInfo->wszProfileName.c_str(),
								FALSE,
								uiServiceIndex,
								(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName.c_str(),
								(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress.c_str()), L"Calling HrAddDelegateMailboxModern.");
					}
				}
				uiServiceIndex++;
			}
			else if (profileInfo->profileServices[i].serviceType == SERVICETYPE_DATAFILE)
			{
				Logger::Write(LOGLEVEL_INFO, L"Adding PST file: " + profileInfo->profileServices[i].pstInfo->wszPstPath);
				CHK_HR_DBG(HrCreatePstService(lpServiceAdmin,
					&lpServiceUid,
					(LPWSTR)profileInfo->profileServices[i].wszServiceName.c_str(),
					profileInfo->profileServices[i].ulResourceFlags,
					profileInfo->profileServices[i].pstInfo->ulPstConfigFlags,
					(LPWSTR)profileInfo->profileServices[i].pstInfo->wszPstPath.c_str(),
					(LPWSTR)profileInfo->profileServices[i].pstInfo->wszDisplayName.c_str()), L"Calling HrCreatePstService.");
				uiServiceIndex++;
			}

		}

		Logger::Write(LOGLEVEL_INFO, L"Setting profile as default.");
		CHK_HR_DBG(HrSetDefaultProfile((LPWSTR)profileInfo->wszProfileName.c_str()), L"Calling HrSetDefaultProfile.");
	}

Error:
	goto CleanUp;
CleanUp:
	return hRes;
}

// Outlook 2013
HRESULT HrSimpleCloneProfile(ProfileInfo * profileInfo, bool bSetDefaultProfile)
{
	HRESULT hRes = S_OK;
	LPSERVICEADMIN2 lpServiceAdmin = NULL;
	unsigned int uiServiceIndex = 0;
	profileInfo->wszProfileName = profileInfo->wszProfileName + L"_Clone";
	Logger::Write(LOGLEVEL_INFO, L"Creating new profile named: " + profileInfo->wszProfileName);
	CHK_HR_DBG(HrCreateProfile((LPWSTR)profileInfo->wszProfileName.c_str(), &lpServiceAdmin), L"Calling HrCreateProfile.");
	if (lpServiceAdmin)
	{
		for (unsigned int i = 0; i < profileInfo->ulServiceCount; i++)
		{
			MAPIUID uidService = { 0 };
			LPMAPIUID lpServiceUid = &uidService;
			if (profileInfo->profileServices[i].serviceType == SERVICETYPE_EXCHANGEACCOUNT)
			{
				Logger::Write(LOGLEVEL_INFO, L"Adding exchange mailbox: " + profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress);
				
				CHK_HR_DBG(HrCreateMsemsServiceMOH(false,
					(LPWSTR)profileInfo->wszProfileName.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszMailboxDN.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN.c_str(),
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerName.c_str(),
					NULL,
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszMailStoreExternalUrl.c_str(),
					NULL,
					(LPWSTR)profileInfo->profileServices[i].exchangeAccountInfo->wszAddressBookExternalUrl.c_str()), L"HrCreateMsemsServiceMOH");

				uiServiceIndex++;
			}
		}
		if (bSetDefaultProfile)
		{
			Logger::Write(LOGLEVEL_INFO, L"Setting profile as default.");
			CHK_HR_DBG(HrSetDefaultProfile((LPWSTR)profileInfo->wszProfileName.c_str()), L"Calling HrSetDefaultProfile.");
		}
	}
Error:
	goto CleanUp;
CleanUp:
	return hRes;
}

VOID PrintProfile(ProfileInfo * profileInfo)
{
	if (profileInfo)
	{
		wprintf(L"Profile name: %ls\n", profileInfo->wszProfileName.c_str());
		wprintf(L"Service count: %i\n", profileInfo->ulServiceCount);
		for (unsigned int i = 0; i < profileInfo->ulServiceCount; i++)
		{
			wprintf(L" -> Service #%i\n", i);
			wprintf(L" -> [%i] Display name: %ls\n", i, profileInfo->profileServices[i].wszDisplayName.c_str());
			wprintf(L" -> [%i] Service name: %ls\n", i, profileInfo->profileServices[i].wszServiceName.c_str());
			wprintf(L" -> [%i] Service resource flags: %#x\n", i, profileInfo->profileServices[i].ulResourceFlags);
			MAPIUID uidService = { 0 };
			LPMAPIUID lpServiceUid = &uidService;
			if (profileInfo->profileServices[i].serviceType == SERVICETYPE_EXCHANGEACCOUNT)
			{
				wprintf(L" -> [%i] Service type: %ls\n", i, L"Exchange Mailbox");
				wprintf(L" -> [%i] E-mail address: %ls\n", i, profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress.c_str());
				wprintf(L" -> [%i] User display name: %ls\n", i, profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName.c_str());
				wprintf(L" -> [%i] OST path: %ls\n", i, profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath.c_str());
				wprintf(L" -> [%i] Config flags: %#x\n", i, profileInfo->profileServices[i].exchangeAccountInfo->ulProfileConfigFlags);
				wprintf(L" -> [%i] Cached mode months: %i\n", i, profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths);
				wprintf(L" -> [%i] Mailbox count: %i\n", i, profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount);
				for (unsigned int j = 0; j < profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount; j++)
				{
					wprintf(L" -> [%i] -> Mailbox #%i\n", i, j);
					if (profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType == PROFILE_DELEGATE)
					{
						wprintf(L" -> [%i] -> [%i] (Delegate) -> E-mail address: %ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress.c_str());
						wprintf(L" -> [%i] -> [%i] (Delegate) -> User display name: %ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName.c_str());
					}
					else
					{
						wprintf(L" -> [%i] -> [%i] (Other)-> E-mail address:%ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress.c_str());
						wprintf(L" -> [%i] -> [%i] (Other) -> User display name:%ls\n", i, j, profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName.c_str());
					}
				}
			}
			else if (profileInfo->profileServices[i].serviceType == SERVICETYPE_DATAFILE)
			{
				wprintf(L" -> [%i] Service type: %ls\n", i, L"PST");
				wprintf(L" -> [%i] Display name: %ls\n", i, profileInfo->profileServices[i].pstInfo->wszDisplayName.c_str());
				wprintf(L" -> [%i] PST path: %ls\n", i, profileInfo->profileServices[i].pstInfo->wszPstPath.c_str());
				wprintf(L" -> [%i] Config flags: %#x\n", i, profileInfo->profileServices[i].pstInfo->ulPstConfigFlags);
			}
		}
	}
Error:
	goto CleanUp;
CleanUp:
	return;
}

HRESULT HrGetProfile(LPWSTR lpszProfileName, ProfileInfo * profileInfo)
{
	HRESULT hRes = S_OK;
	profileInfo->wszProfileName = ConvertWideCharToStdWstring(lpszProfileName);

	LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
	LPSRestriction lpProfRes = NULL;
	LPSRestriction lpProfResLvl1 = NULL;
	LPSPropValue lpProfPropVal = NULL;
	LPMAPITABLE lpProfTable = NULL;
	LPSRowSet lpProfRows = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPMAPITABLE lpServiceTable = NULL;
	// Setting up an enum and a prop tag array with the props we'll use
	enum { iDisplayName, iDefaultProfile, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME, PR_DEFAULT_PROFILE };

	CHK_HR_DBG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles."); // Pointer to new IProfAdmin
													   // Get an IProfAdmin interface.

	CHK_HR_DBG(lpProfAdmin->GetProfileTable(0,
		&lpProfTable), L"Calling GetProfileTable.");

	// Allocate memory for the restriction
	CHK_HR_DBG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)&lpProfRes), L"Calling MAPIAllocateBuffer.");

	CHK_HR_DBG(MAPIAllocateBuffer(
		sizeof(SRestriction) * 2,
		(LPVOID*)&lpProfResLvl1), L"Calling MAPIAllocateBuffer");

	CHK_HR_DBG(MAPIAllocateBuffer(
		sizeof(SPropValue),
		(LPVOID*)&lpProfPropVal), L"Calling MAPIAllocateBuffer");

	// Set up restriction to query the profile table
	lpProfRes->rt = RES_AND;
	lpProfRes->res.resAnd.cRes = 0x00000002;
	lpProfRes->res.resAnd.lpRes = lpProfResLvl1;

	lpProfResLvl1[0].rt = RES_EXIST;
	lpProfResLvl1[0].res.resExist.ulPropTag = PR_DISPLAY_NAME_A;
	lpProfResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
	lpProfResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
	lpProfResLvl1[1].rt = RES_PROPERTY;
	lpProfResLvl1[1].res.resProperty.relop = RELOP_EQ;
	lpProfResLvl1[1].res.resProperty.ulPropTag = PR_DISPLAY_NAME_A;
	lpProfResLvl1[1].res.resProperty.lpProp = lpProfPropVal;

	lpProfPropVal->ulPropTag = PR_DISPLAY_NAME_A;
	lpProfPropVal->Value.lpszA = ConvertWideCharToMultiByte(lpszProfileName);

	// Query the table to get the the default profile only
	CHK_HR_DBG(HrQueryAllRows(lpProfTable,
		(LPSPropTagArray)&sptaProps,
		lpProfRes,
		NULL,
		0,
		&lpProfRows), L"Calling HrQueryAllRows.");

	if (lpProfRows->cRows == 0)
	{
		return MAPI_E_NOT_FOUND;
	}
	else if (lpProfRows->cRows == 1)
	{
		profileInfo->bDefaultProfile = lpProfRows->aRow->lpProps[iDefaultProfile].Value.b;
	}
	else
	{
		return MAPI_E_CALL_FAILED;
	}

	// Begin process services

	CHK_HR_DBG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		MAPI_UNICODE,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices.");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		lpServiceAdmin->GetMsgServiceTable(0,
			&lpServiceTable);
		LPSRestriction lpSvcRes = NULL;
		LPSRestriction lpsvcResLvl1 = NULL;
		LPSPropValue lpSvcPropVal = NULL;
		LPSRowSet lpSvcRows = NULL;

		// Setting up an enum and a prop tag array with the props we'll use
		enum { iServiceUid, iServiceName, iDisplayNameW, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
		SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID ,PR_SERVICE_NAME_A, PR_DISPLAY_NAME_W, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

		// Query the table to get the the default profile only
		CHK_HR_DBG(HrQueryAllRows(lpServiceTable,
			(LPSPropTagArray)&sptaSvcProps,
			NULL,
			NULL,
			0,
			&lpSvcRows), L"Calling HrQueryAllRows.");

		if (lpSvcRows->cRows > 0)
		{
			profileInfo->ulServiceCount = lpSvcRows->cRows;
			profileInfo->profileServices = new ServiceInfo[lpSvcRows->cRows];


			// Start loop services
			for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
			{
				ZeroMemory(&profileInfo->profileServices[i], sizeof(ServiceInfo));
				profileInfo->profileServices[i].wszServiceName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA));
				profileInfo->profileServices[i].ulResourceFlags = lpSvcRows->aRow[i].lpProps[iServiceResFlags].Value.l;
				profileInfo->profileServices[i].wszDisplayName = lpSvcRows->aRow[i].lpProps[iDisplayNameW].Value.lpszW;
				profileInfo->profileServices[i].serviceType = SERVICETYPE_UNKNOWN;;
				if (profileInfo->profileServices[i].ulResourceFlags & SERVICE_DEFAULT_STORE)
				{
					profileInfo->profileServices[i].bDefaultStore = true;
				}
				// Exchange account
				if (0 == strcmp(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA, "MSEMS"))
				{
					profileInfo->profileServices[i].exchangeAccountInfo = new ExchangeAccountInfo();
					profileInfo->profileServices[i].serviceType = SERVICETYPE_EXCHANGEACCOUNT;
					LPPROVIDERADMIN lpProvAdmin = NULL;

					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{

						// Read the EMSMDB section
						LPPROFSECT lpProfSect = NULL;
						if (SUCCEEDED(lpProvAdmin->OpenProfileSection((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iEmsMdbSectUid].Value.bin.lpb,
							NULL,
							0L,
							&lpProfSect)))
						{
							LPMAPIPROP pMAPIProp = NULL;
							if (SUCCEEDED(lpProfSect->QueryInterface(IID_IMAPIProp, (void**)&pMAPIProp)))
							{
								// bind to the PR_RULE_ACTION_TYPE property to get the ammount to sync
								LPSPropValue profilePrRuleActionType = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_RULE_ACTION_TYPE, &profilePrRuleActionType)))
								{
									if (profilePrRuleActionType)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths = profilePrRuleActionType->Value.i;
										if (profilePrRuleActionType) MAPIFreeBuffer(profilePrRuleActionType);
									}

								}
								else
								{
									profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths = 0;
								}

								// PR_PROFILE_OFFLINE_STORE_PATH_W
								LPSPropValue profileOfflineStorePath = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_OFFLINE_STORE_PATH_W, &profileOfflineStorePath)))
								{
									if (profileOfflineStorePath)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = profileOfflineStorePath->Value.lpszW;
										if (profileOfflineStorePath) MAPIFreeBuffer(profileOfflineStorePath);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = L"";
									}

								}
								else
								{
									profileInfo->profileServices[i].exchangeAccountInfo->iCachedModeMonths = 0;
								}
								// bind to the PR_PROFILE_CONFIG_FLAGS property
								LPSPropValue profileConfigFlags = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_CONFIG_FLAGS, &profileConfigFlags)))
								{
									if (profileConfigFlags)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->ulProfileConfigFlags = profileConfigFlags->Value.l;
									}
									MAPIFreeBuffer(profileConfigFlags);
								}
								// bind to the PR_PROFILE_USER_SMTP_EMAIL_ADDRESS property
								LPSPropValue profileUserSmtpEmailAddress = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_USER_SMTP_EMAIL_ADDRESS, &profileUserSmtpEmailAddress)))
								{
									if (profileUserSmtpEmailAddress)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress = SubstringToEnd(L"smtp:", ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUserSmtpEmailAddress->Value.lpszA)));
										if (profileUserSmtpEmailAddress) MAPIFreeBuffer(profileUserSmtpEmailAddress);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszEmailAddress = std::wstring(L" ");
									}
								}
								LPSPropValue profileDisplayName = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_DISPLAY_NAME_A, &profileDisplayName)))
								{
									if (profileDisplayName)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileDisplayName->Value.lpszA));
										if (profileDisplayName) MAPIFreeBuffer(profileDisplayName);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDisplayName = std::wstring(L" ");
									}
								}

								// PR_PROFILE_USER
								LPSPropValue profileUser = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_USER, &profileUser)))
								{
									if (profileUser)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszMailboxDN = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUser->Value.lpszA));
										if (profileUser) MAPIFreeBuffer(profileUser);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszMailboxDN = std::wstring(L" ");
									}
								}

								// PR_PROFILE_HOME_SERVER_DN
								LPSPropValue profileHomeServerDN = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_HOME_SERVER_DN, &profileHomeServerDN)))
								{
									if (profileHomeServerDN)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileHomeServerDN->Value.lpszA));
										if (profileHomeServerDN) MAPIFreeBuffer(profileHomeServerDN);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = std::wstring(L" ");
									}
								}


								// PR_PROFILE_HOME_SERVER
								LPSPropValue profileHomeServer = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_HOME_SERVER, &profileHomeServer)))
								{
									if (profileHomeServer)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerName = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileHomeServer->Value.lpszA));
										if (profileHomeServer) MAPIFreeBuffer(profileHomeServer);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszHomeServerDN = std::wstring(L" ");
									}
								}

								// PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL
								LPSPropValue profileMapiHttpMailStoreExternal = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL, &profileMapiHttpMailStoreExternal)))
								{
									if (profileMapiHttpMailStoreExternal)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszMailStoreExternalUrl = profileMapiHttpMailStoreExternal->Value.lpszW;
										if (profileMapiHttpMailStoreExternal) MAPIFreeBuffer(profileMapiHttpMailStoreExternal);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = L"";
									}

								}

								// PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL
								LPSPropValue profileMapiHttpAddressbookExternal = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_PROFILE_MAPIHTTP_ADDRESSBOOK_EXTERNAL_URL, &profileMapiHttpAddressbookExternal)))
								{
									if (profileMapiHttpAddressbookExternal)
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszAddressBookExternalUrl = profileMapiHttpAddressbookExternal->Value.lpszW;
										if (profileMapiHttpAddressbookExternal) MAPIFreeBuffer(profileMapiHttpAddressbookExternal);
									}
									else
									{
										profileInfo->profileServices[i].exchangeAccountInfo->wszDatafilePath = L"";
									}
								}

								// PR_SERVICE_UID
								LPSPropValue serviceUid = NULL;
								if (SUCCEEDED(HrGetOneProp(pMAPIProp, PR_SERVICE_UID, &serviceUid)))
								{
									if (serviceUid)
									{
										LPMAPIUID lpMuidServiceUid = NULL;
										lpMuidServiceUid = &profileInfo->profileServices[i].muidServiceUid;
										memcpy(lpMuidServiceUid, serviceUid->Value.bin.lpb, sizeof(MAPIUID));
										if (serviceUid) MAPIFreeBuffer(serviceUid);
									}
								}
							}
							if (lpProfSect) lpProfSect->Release();
						}

						// End read the EMSMDB section


						// Loop providers
						LPMAPITABLE lpProvTable = NULL;
						LPSRestriction lpProvRes = NULL;
						LPSRestriction lpProvResLvl1 = NULL;
						LPSPropValue lpProvPropVal = NULL;
						LPSRowSet lpProvRows = NULL;

						// Setting up an enum and a prop tag array with the props we'll use
						enum { iProvInstanceKey, cptaProvProps };
						SizedSPropTagArray(cptaProvProps, sptaProvProps) = { cptaProvProps, PR_INSTANCE_KEY };

						// Allocate memory for the restriction
						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SRestriction),
							(LPVOID*)&lpProvRes), L"Calling MAPIAllocateBuffer.");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)&lpProvResLvl1), L"Calling MAPIAllocateBuffer.");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)&lpProvPropVal), L"Calling MAPIAllocateBuffer.");

						// Set up restriction to query the provider table
						lpProvRes->rt = RES_AND;
						lpProvRes->res.resAnd.cRes = 0x00000002;
						lpProvRes->res.resAnd.lpRes = lpProvResLvl1;

						lpProvResLvl1[0].rt = RES_EXIST;
						lpProvResLvl1[0].res.resExist.ulPropTag = PR_RESOURCE_TYPE;
						lpProvResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
						lpProvResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
						lpProvResLvl1[1].rt = RES_PROPERTY;
						lpProvResLvl1[1].res.resProperty.ulPropTag = PR_RESOURCE_TYPE;
						lpProvResLvl1[1].res.resProperty.lpProp = lpProvPropVal;
						lpProvResLvl1[1].res.resProperty.relop = RELOP_EQ;

						lpProvPropVal->ulPropTag = PR_RESOURCE_TYPE;
						lpProvPropVal->Value.l = MAPI_STORE_PROVIDER;

						lpProvAdmin->GetProviderTable(0,
							&lpProvTable);
						// Query the table to get the the default profile only
						CHK_HR_DBG(HrQueryAllRows(lpProvTable,
							(LPSPropTagArray)&sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), L"Calling HrQueryAllRows.");

						if (lpProvRows->cRows > 0)
						{
							profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount = lpProvRows->cRows;
							profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes = new MailboxInfo[lpProvRows->cRows];

							for (unsigned int j = 0; j < lpProvRows->cRows; j++)
							{
								ZeroMemory(&profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j], sizeof(MailboxInfo));
								profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName = std::wstring(L" ");
								LPPROFSECT lpProfSection = NULL;
								if (SUCCEEDED(lpServiceAdmin->OpenProfileSection((LPMAPIUID)lpProvRows->aRow[j].lpProps[iProvInstanceKey].Value.bin.lpb, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSection)))
								{

									LPMAPIPROP lpMAPIProp = NULL;
									if (SUCCEEDED(lpProfSection->QueryInterface(IID_IMAPIProp, (void**)&lpMAPIProp)))
									{
										// PR_DISPLAY_NAME
										LPSPropValue prDisplayName = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_DISPLAY_NAME, &prDisplayName)))
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName = ConvertWideCharToStdWstring(prDisplayName->Value.lpszW);
											if (prDisplayName) MAPIFreeBuffer(prDisplayName);
										}
										else
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszDisplayName = std::wstring(L" ");
										}

										// PR_PROFILE_TYPE
										LPSPropValue prProfileType = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_TYPE, &prProfileType)))
										{
											profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulProfileType = prProfileType->Value.l;
										}

										// PR_PROFILE_USER_SMTP_EMAIL_ADDRESS
										LPSPropValue profileUserSmtpEmailAddress = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_USER_SMTP_EMAIL_ADDRESS, &profileUserSmtpEmailAddress)))
										{
											if (profileUserSmtpEmailAddress)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileUserSmtpEmailAddress->Value.lpszA));
												if (profileUserSmtpEmailAddress) MAPIFreeBuffer(profileUserSmtpEmailAddress);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszSmtpAddress = std::wstring(L" ");
											}
										}

										// PR_PROFILE_MAILBOX
										LPSPropValue profileMailbox = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_MAILBOX, &profileMailbox)))
										{
											if (profileMailbox)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileMailbox = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileMailbox->Value.lpszA));
												if (profileMailbox) MAPIFreeBuffer(profileMailbox);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileMailbox = std::wstring(L" ");
											}
										}

										// PR_PROFILE_SERVER_DN
										LPSPropValue profileServerDN = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_SERVER_DN, &profileServerDN)))
										{
											if (profileMailbox)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerDN = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileServerDN->Value.lpszA));
												if (profileServerDN) MAPIFreeBuffer(profileServerDN);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerDN = std::wstring(L" ");
											}
										}

										// PR_ROH_PROXY_SERVER
										LPSPropValue rohProxyServer = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_ROH_PROXY_SERVER, &rohProxyServer)))
										{
											if (rohProxyServer)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszRohProxyServer = ConvertWideCharToStdWstring(rohProxyServer->Value.lpszW);
												if (rohProxyServer) MAPIFreeBuffer(rohProxyServer);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszRohProxyServer = std::wstring(L" ");
											}
										}

										// PR_PROFILE_SERVER
										LPSPropValue profileServer = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_SERVER, &profileServer)))
										{
											if (profileServer)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServer = ConvertWideCharToStdWstring(ConvertMultiByteToWideChar(profileServer->Value.lpszA));
												if (profileServer) MAPIFreeBuffer(profileServer);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServer = std::wstring(L" ");
											}
										}

										// PR_PROFILE_SERVER_FQDN_W
										LPSPropValue profileServerFqdnW = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_SERVER_FQDN_W, &profileServerFqdnW)))
										{
											if (profileServerFqdnW)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerFqdnW = ConvertWideCharToStdWstring(profileServerFqdnW->Value.lpszW);
												if (profileServerFqdnW) MAPIFreeBuffer(profileServerFqdnW);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszProfileServerFqdnW = std::wstring(L" ");
											}
										}

										// PR_PROFILE_LKG_AUTODISCOVER_URL
										LPSPropValue profileAutodiscoverUrl = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_LKG_AUTODISCOVER_URL, &profileAutodiscoverUrl)))
										{
											if (profileAutodiscoverUrl)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszAutodiscoverUrl = ConvertWideCharToStdWstring(profileAutodiscoverUrl->Value.lpszW);
												if (profileServerFqdnW) MAPIFreeBuffer(profileAutodiscoverUrl);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszAutodiscoverUrl = std::wstring(L" ");
											}
										}

										// PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL
										LPSPropValue profileMailStoreInternalUrl = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_MAPIHTTP_MAILSTORE_INTERNAL_URL, &profileMailStoreInternalUrl)))
										{
											if (profileMailStoreInternalUrl)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreInternalUrl = ConvertWideCharToStdWstring(profileMailStoreInternalUrl->Value.lpszW);
												if (profileMailStoreInternalUrl) MAPIFreeBuffer(profileMailStoreInternalUrl);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreInternalUrl = std::wstring(L" ");
											}
										}

										// PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL
										LPSPropValue profileMailStoreExternalUrl = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_MAPIHTTP_MAILSTORE_EXTERNAL_URL, &profileMailStoreExternalUrl)))
										{
											if (profileMailStoreExternalUrl)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreExternalUrl = ConvertWideCharToStdWstring(profileMailStoreExternalUrl->Value.lpszW);
												if (profileMailStoreExternalUrl) MAPIFreeBuffer(profileMailStoreExternalUrl);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].wszMailStoreExternalUrl = std::wstring(L" ");
											}
										}

										// PR_RESOURCE_FLAGS
										LPSPropValue resourceFlags = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_RESOURCE_FLAGS, &resourceFlags)))
										{
											if (resourceFlags)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulResourceFlags = resourceFlags->Value.l;
												if (resourceFlags) MAPIFreeBuffer(resourceFlags);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulResourceFlags = 0;
											}
										}

										// PR_PROFILE_RPC_PROXY_SERVER_AUTH_PACKAGE
										LPSPropValue rohAuthPackage = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_RPC_PROXY_SERVER_AUTH_PACKAGE, &rohAuthPackage)))
										{
											if (rohAuthPackage)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohProxyAuthScheme = rohAuthPackage->Value.l;
												if (rohAuthPackage) MAPIFreeBuffer(rohAuthPackage);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohProxyAuthScheme = 0;
											}
										}

										// PR_ROH_FLAGS
										LPSPropValue rohFlags = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_ROH_FLAGS, &rohFlags)))
										{
											if (rohFlags)
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohFlags = rohFlags->Value.l;
												if (rohFlags) MAPIFreeBuffer(rohFlags);
											}
											else
											{
												profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].ulRohFlags = 0;
											}
										}

										// PR_SERVICE_UID
										LPSPropValue serviceUid = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_SERVICE_UID, &serviceUid)))
										{
											if (serviceUid)
											{
												LPMAPIUID lpMuidServiceUid = NULL;
												lpMuidServiceUid = &profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].muidServiceUid;
												memcpy(lpMuidServiceUid, (LPMAPIUID)serviceUid->Value.bin.lpb, sizeof(MAPIUID));
												if (serviceUid) MAPIFreeBuffer(serviceUid);
											}
										}
										
										// PR_PROFILE_ALTERNATE_STORE_TYPE
										LPSPropValue alternateStoreType = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROFILE_ALTERNATE_STORE_TYPE, &alternateStoreType)))
										{
											if (alternateStoreType)
											{
												if (ConvertWideCharToStdWstring(alternateStoreType->Value.lpszW) == L"Archive")
												{
													profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive = true;
												}
												else
												{
													profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].bIsOnlineArchive = false;
												}
												if (alternateStoreType) MAPIFreeBuffer(alternateStoreType);
											}
										}

										// PR_PROVIDER_UID
										LPMAPIUID lpMuidProviderUid = NULL;
										lpMuidProviderUid = &profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].muidProviderUid;
										memcpy(lpMuidProviderUid, (LPMAPIUID)lpProvRows->aRow[j].lpProps[iProvInstanceKey].Value.bin.lpb, sizeof(MAPIUID));

										//LPSPropValue providerUid = NULL;
										//hRes = HrGetOneProp(lpMAPIProp, PR_INSTANCE_KEY, &providerUid);
										//if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PROVIDER_UID, &providerUid)))
										//{
										//	if (providerUid)
										//	{
										//		LPMAPIUID lpMuidProviderUid = NULL;
										//		lpMuidProviderUid = &profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes[j].muidProviderUid;
										//		memcpy(lpMuidProviderUid, (LPMAPIUID)providerUid->Value.bin.lpb, sizeof(MAPIUID));
										//		if (providerUid) MAPIFreeBuffer(providerUid);
										//	}
										//}


									}
								}
							}
							if (lpProvRows) FreeProws(lpProvRows);
						}
						else
						{
							profileInfo->profileServices[i].exchangeAccountInfo->ulMailboxCount = lpProvRows->cRows;
							profileInfo->profileServices[i].exchangeAccountInfo->accountMailboxes = new MailboxInfo();
						}
						if (lpProvPropVal) MAPIFreeBuffer(lpProvPropVal);
						if (lpProvResLvl1) MAPIFreeBuffer(lpProvResLvl1);
						if (lpProvRes) MAPIFreeBuffer(lpProvRes);
						if (lpProvTable) lpProvTable->Release();
						//End Loop Providers
						if (lpProvAdmin) lpProvAdmin->Release();
					}

				}

				else if ((0 == strcmp(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA, "MSPST MS")) || (0 == strcmp(lpSvcRows->aRow[i].lpProps[iServiceName].Value.lpszA, "MSUPST MS")))
				{
					profileInfo->profileServices[i].pstInfo = new PstInfo();
					profileInfo->profileServices[i].pstInfo->wszDisplayName = std::wstring(L" ");
					profileInfo->profileServices[i].pstInfo->wszPstPath = std::wstring(L" ");
					profileInfo->profileServices[i].pstInfo->ulPstConfigFlags = 0;
					profileInfo->profileServices[i].serviceType = SERVICETYPE_DATAFILE;

					LPPROVIDERADMIN lpProvAdmin = NULL;

					if (SUCCEEDED(lpServiceAdmin->AdminProviders((LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb,
						0,
						&lpProvAdmin)))
					{
						// Loop providers
						LPMAPITABLE lpProvTable = NULL;
						LPSRestriction lpProvRes = NULL;
						LPSRestriction lpProvResLvl1 = NULL;
						LPSPropValue lpProvPropVal = NULL;
						LPSRowSet lpProvRows = NULL;

						// Setting up an enum and a prop tag array with the props we'll use
						enum { iProvInstanceKey, cptaProvProps };
						SizedSPropTagArray(cptaProvProps, sptaProvProps) = { cptaProvProps, PR_INSTANCE_KEY };

						// Allocate memory for the restriction
						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SRestriction),
							(LPVOID*)&lpProvRes), L"Calling MAPIAllocateBuffer");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)&lpProvResLvl1), L"Calling MAPIAllocateBuffer");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)&lpProvPropVal), L"Calling MAPIAllocateBuffer");

						// Set up restriction to query the provider table
						lpProvRes->rt = RES_AND;
						lpProvRes->res.resAnd.cRes = 0x00000002;
						lpProvRes->res.resAnd.lpRes = lpProvResLvl1;

						lpProvResLvl1[0].rt = RES_EXIST;
						lpProvResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_UID;
						lpProvResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
						lpProvResLvl1[0].res.resExist.ulReserved2 = 0x00000000;
						lpProvResLvl1[1].rt = RES_PROPERTY;
						lpProvResLvl1[1].res.resProperty.relop = RELOP_EQ;
						lpProvResLvl1[1].res.resProperty.ulPropTag = PR_SERVICE_UID;
						lpProvResLvl1[1].res.resProperty.lpProp = lpProvPropVal;

						lpProvPropVal->ulPropTag = PR_SERVICE_UID;
						lpProvPropVal->Value = lpSvcRows->aRow[i].lpProps[iServiceUid].Value;

						lpProvAdmin->GetProviderTable(0,
							&lpProvTable);
						// Query the table to get the the default profile only
						CHK_HR_DBG(HrQueryAllRows(lpProvTable,
							(LPSPropTagArray)&sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), L"HrGetProfile");

						if (lpProvRows->cRows > 0)
						{

							LPPROFSECT lpProfSection = NULL;
							if (SUCCEEDED(lpServiceAdmin->OpenProfileSection((LPMAPIUID)lpProvRows->aRow->lpProps[iProvInstanceKey].Value.bin.lpb, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSection)))
							{
								LPMAPIPROP lpMAPIProp = NULL;
								if (SUCCEEDED(lpProfSection->QueryInterface(IID_IMAPIProp, (void**)&lpMAPIProp)))
								{
									LPSPropValue prDisplayName = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_DISPLAY_NAME_W, &prDisplayName)))
									{
										profileInfo->profileServices[i].pstInfo->wszDisplayName = ConvertWideCharToStdWstring(prDisplayName->Value.lpszW);
										if (prDisplayName) MAPIFreeBuffer(prDisplayName);
									}
									else
									{
										profileInfo->profileServices[i].pstInfo->wszDisplayName = std::wstring(L" ");
									}
									// bind to the PR_PST_PATH_W property
									LPSPropValue pstPathW = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_PATH_W, &pstPathW)))
									{
										if (pstPathW)
										{
											profileInfo->profileServices[i].pstInfo->wszPstPath = ConvertWideCharToStdWstring(pstPathW->Value.lpszW);
											if (pstPathW) MAPIFreeBuffer(pstPathW);
										}
										else
										{
											profileInfo->profileServices[i].pstInfo->wszPstPath = std::wstring(L" ");
										}
									}
									// bind to the PR_PST_CONFIG_FLAGS property to get the ammount to sync
									LPSPropValue pstConfigFlags = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_CONFIG_FLAGS, &pstConfigFlags)))
									{
										if (pstConfigFlags)
										{
											profileInfo->profileServices[i].pstInfo->ulPstConfigFlags = pstConfigFlags->Value.l;
											if (pstConfigFlags) MAPIFreeBuffer(pstConfigFlags);
										}
									}
								}
							}

							if (lpProvRows) FreeProws(lpProvRows);
						}
						if (lpProvPropVal) MAPIFreeBuffer(lpProvPropVal);
						if (lpProvResLvl1) MAPIFreeBuffer(lpProvResLvl1);
						if (lpProvRes) MAPIFreeBuffer(lpProvRes);
						if (lpProvTable) lpProvTable->Release();
						//End Loop Providers
						if (lpProvAdmin) lpProvAdmin->Release();
					}

				}

			}
			if (lpSvcRows) FreeProws(lpSvcRows);
			// End loop services


		}

		if (lpSvcPropVal) MAPIFreeBuffer(lpSvcPropVal);
		if (lpsvcResLvl1) MAPIFreeBuffer(lpsvcResLvl1);
		if (lpSvcRes) MAPIFreeBuffer(lpSvcRes);
		if (lpServiceTable) lpServiceTable->Release();
		if (lpServiceAdmin) lpServiceAdmin->Release();

	}
	// End process services
Error:
	goto CleanUp;
CleanUp:
	// Free up memory
	if (lpProfRows) FreeProws(lpProfRows);
	if (lpProfTable) lpProfTable->Release();
	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;

}

// HrDeleteProvider
// Deletes the provider with the specified UID from the service with the specified UID in a given profile
HRESULT HrDeleteProvider(LPWSTR lpwszProfileName, LPMAPIUID lpServiceUid, LPMAPIUID lpProviderUid)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;
	LPPROVIDERADMIN lpProviderAdmin = NULL;

	CHK_HR_DBG(MAPIAdminProfiles(0, // Flags
		&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin

	CHK_HR_DBG(lpProfAdmin->AdminServices((LPTSTR)ConvertWideCharToMultiByte(lpwszProfileName),
		LPTSTR(""),            // Password for that profile.
		NULL,                // Handle to parent window.
		0,                    // Flags.
		&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

	if (lpServiceAdmin)
	{
		CHK_HR_DBG(lpServiceAdmin->AdminProviders(lpServiceUid, NULL, &lpProviderAdmin), L"Calling AdminProviders");
		if (lpProviderAdmin)
		{
			CHK_HR_DBG(lpProviderAdmin->DeleteProvider(lpProviderUid), L"Calling DeleteProvider");
		}
	}

Error:
	goto CleanUp;
CleanUp:
	return hRes;
}

// HrGetSections
// Returns the EMSMDB and StoreProvider sections of a service
HRESULT HrGetSections(LPSERVICEADMIN2 lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT* lppEmsMdbSection, LPPROFSECT* lppStoreProviderSection)
{
	HRESULT hRes = S_OK;
	SizedSPropTagArray(2, sptaUids) = { 2,{ PR_STORE_PROVIDERS, PR_EMSMDB_SECTION_UID } };
	ULONG				cValues = 0;
	LPSPropValue		lpProps = nullptr;
	LPPROFSECT			lpSvcProfSect = nullptr;
	LPPROFSECT			lpEmsMdbProfSect = nullptr;
	LPPROFSECT			lpStoreProvProfSect = nullptr;

	if (!lpSvcAdmin || !lpServiceUid || !lppStoreProviderSection)
		return E_INVALIDARG;

	if (NULL != lppStoreProviderSection)
	{
		*lppStoreProviderSection = nullptr;
	}
	if (NULL != lppEmsMdbSection)
	{
		*lppEmsMdbSection = nullptr;
	}

	CHK_HR_DBG(lpSvcAdmin->OpenProfileSection(lpServiceUid, NULL, MAPI_FORCE_ACCESS | MAPI_MODIFY, &lpSvcProfSect), L"Calling OpenProfileSection.");

	CHK_HR_DBG(lpSvcProfSect->GetProps(
		(LPSPropTagArray)& sptaUids,
		0,
		&cValues,
		&lpProps), L"Calling GetProps.");

	if (cValues != 2)
		return E_FAIL;


	if (lpProps[0].ulPropTag != sptaUids.aulPropTag[0])
		CHK_HR_DBG(lpProps[0].Value.err, L"Cheking Value.err");
	if (NULL != lppEmsMdbSection)
	{
		if (lpProps[1].ulPropTag != sptaUids.aulPropTag[1])
			CHK_HR_DBG(lpProps[1].Value.err, L"Cheking Value.err");
	}

	if (NULL != lppStoreProviderSection)
	{
		CHK_HR_DBG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[0].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpStoreProvProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
	{
		CHK_HR_DBG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[1].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpEmsMdbProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
		* lppEmsMdbSection = lpEmsMdbProfSect;

	if (NULL != lppStoreProviderSection)
		* lppStoreProviderSection = lpStoreProvProfSect;

	if (lpSvcProfSect) lpSvcProfSect->Release();
	if (lpProps)
	{
		MAPIFreeBuffer(lpProps);
		lpProps = nullptr;
	}

Error:
	goto CleanUp;
CleanUp:
	return hRes;
}

// HrGetSections
// Returns the EMSMDB and StoreProvider sections of a service
HRESULT HrGetSections(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID lpServiceUid, LPPROFSECT * lppEmsMdbSection, LPPROFSECT * lppStoreProviderSection)
{
	HRESULT hRes = S_OK;
	SizedSPropTagArray(2, sptaUids) = { 2,{ PR_STORE_PROVIDERS, PR_EMSMDB_SECTION_UID } };
	ULONG				cValues = 0;
	LPSPropValue		lpProps = nullptr;
	LPPROFSECT			lpSvcProfSect = nullptr;
	LPPROFSECT			lpEmsMdbProfSect = nullptr;
	LPPROFSECT			lpStoreProvProfSect = nullptr;

	if (!lpSvcAdmin || !lpServiceUid || !lppStoreProviderSection)
		return E_INVALIDARG;

	if (NULL != lppStoreProviderSection)
	{
		*lppStoreProviderSection = nullptr;
	}
	if (NULL != lppEmsMdbSection)
	{
		*lppEmsMdbSection = nullptr;
	}

	CHK_HR_DBG(lpSvcAdmin->OpenProfileSection(lpServiceUid, NULL, MAPI_FORCE_ACCESS | MAPI_MODIFY, &lpSvcProfSect), L"Calling OpenProfileSection.");

	CHK_HR_DBG(lpSvcProfSect->GetProps(
		(LPSPropTagArray)& sptaUids,
		0,
		&cValues,
		&lpProps), L"Calling GetProps.");

	if (cValues != 2)
		return E_FAIL;


	if (lpProps[0].ulPropTag != sptaUids.aulPropTag[0])
		CHK_HR_DBG(lpProps[0].Value.err, L"Cheking Value.err");
	if (NULL != lppEmsMdbSection)
	{
		if (lpProps[1].ulPropTag != sptaUids.aulPropTag[1])
			CHK_HR_DBG(lpProps[1].Value.err, L"Cheking Value.err");
	}

	if (NULL != lpStoreProvProfSect)
	{
		CHK_HR_DBG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[0].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpStoreProvProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
	{
		CHK_HR_DBG(lpSvcAdmin->OpenProfileSection(
			(LPMAPIUID)lpProps[1].Value.bin.lpb,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpEmsMdbProfSect), L"Calling OpenProfileSection.");
	}

	if (NULL != lppEmsMdbSection)
		* lppEmsMdbSection = lpEmsMdbProfSect;

	if (NULL != lppStoreProviderSection)
		* lppStoreProviderSection = lpStoreProvProfSect;

	if (lpSvcProfSect) lpSvcProfSect->Release();
	if (lpProps)
	{
		MAPIFreeBuffer(lpProps);
		lpProps = nullptr;
	}

Error:
	goto CleanUp;
CleanUp:
	return hRes;
}

HRESULT ListABService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID pMAPIUid)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE		lpMsgSvcTable = NULL; // MAPI table pointer.
	LPSRowSet		lpSvcRows = NULL;
	LPSRestriction	lpRes = NULL;
	LPSPropValue	lpspvSvcUid = NULL;

	enum { iServiceUid, iDisplayName, iAbServerName, iAbServerPort, iAbUsername, iAbSearchBase, iAbSearchTimeout, iAbMaxEntries, iAbUseSSL, iAbRequireSpa, AbEnableBrowsing, iAbDefaultSearchBase, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_SERVICE_UID, PR_DISPLAY_NAME, PROP_AB_PROVIDER_SERVER_NAME, PROP_AB_PROVIDER_SERVER_PORT,
		PROP_AB_PROVIDER_USER_NAME, PROP_AB_PROVIDER_SEARCH_BASE, PROP_AB_PROVIDER_SEARCH_TIMEOUT, PROP_AB_PROVIDER_MAX_ENTRIES, PROP_AB_PROVIDER_USE_SSL,
		PROP_AB_PROVIDER_SERVER_SPA, PROP_AB_PROVIDER_ENABLE_BROWSING, PROP_AB_PROVIDER_SEARCH_BASE_DEFAULT };

	LPPROFSECT lpProfSect = NULL;
	LPMAPIPROP lpMapiProp = NULL;
	ULONG ulPropVal = 0;
	LPSPropValue lpsPropValues = NULL;
	HCK(lpSvcAdmin2->OpenProfileSection(pMAPIUid, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSect));

	HCK(lpProfSect->GetProps((LPSPropTagArray)& sptaProps, NULL, &ulPropVal, &lpsPropValues));
	if (lpsPropValues)
	{
		Logger::WriteLine(LOGLEVEL_INFO, L"  Display Name        : " + ConvertWideCharToStdWstring(lpsPropValues[iDisplayName].Value.lpszW));
		Logger::WriteLine(LOGLEVEL_INFO, L"  Ldap Server Name    : " + ConvertMultiByteToStdWString(lpsPropValues[iAbServerName].Value.lpszA));
		Logger::WriteLine(LOGLEVEL_INFO, L"  Ldap Server Port    : " + ConvertMultiByteToStdWString(lpsPropValues[iAbServerPort].Value.lpszA));
		Logger::WriteLine(LOGLEVEL_INFO, L"  Username            : " + ConvertMultiByteToStdWString(lpsPropValues[iAbUsername].Value.lpszA));
		Logger::WriteLine(LOGLEVEL_INFO, L"  Search Base         : " + ConvertMultiByteToStdWString(lpsPropValues[iAbSearchBase].Value.lpszA));
		Logger::WriteLine(LOGLEVEL_INFO, L"  Search Timeout      : " + ConvertMultiByteToStdWString(lpsPropValues[iAbSearchTimeout].Value.lpszA));
		Logger::WriteLine(LOGLEVEL_INFO, L"  Maximum entries     : " + ConvertMultiByteToStdWString(lpsPropValues[iAbMaxEntries].Value.lpszA));
		if (lpsPropValues[iAbUseSSL].Value.b)
			Logger::WriteLine(LOGLEVEL_INFO, L"  Use SSL             : true");
		else
			Logger::WriteLine(LOGLEVEL_INFO, L"  Use SSL             : false");
		if (lpsPropValues[iAbRequireSpa].Value.b)
			Logger::WriteLine(LOGLEVEL_INFO, L"  Require SPA         : true");
		else
			Logger::WriteLine(LOGLEVEL_INFO, L"  Require SPA         : false");
		if (lpsPropValues[AbEnableBrowsing].Value.b)
			Logger::WriteLine(LOGLEVEL_INFO, L"  Enable browsing     : true");
		else
			Logger::WriteLine(LOGLEVEL_INFO, L"  Enable browsing     : false");
		if (lpsPropValues[iAbDefaultSearchBase].Value.i == 1)
			Logger::WriteLine(LOGLEVEL_INFO, L"  Default search base : true");
		else
			Logger::WriteLine(LOGLEVEL_INFO, L"  Default search base : false");
	}
	else
		Logger::WriteLine(LOGLEVEL_FAILED, L"Unable to retrieve Ldap AB properties.");

Error:
	return hRes;
}

HRESULT ListAllABServices(LPSERVICEADMIN2 lpSvcAdmin2)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE		lpMsgSvcTable = NULL; // MAPI table pointer.
	LPSRowSet		lpSvcRows = NULL;
	LPSRestriction	lpRes = NULL;
	LPSPropValue	lpspvSvcName = NULL;

	enum { iServiceUid, iDisplayName, iAbServerName, iAbServerPort, iAbUsername, iAbSearchBase, iAbSearchTimeout, iAbMaxEntries, iAbUseSSL, iAbRequireSpa, AbEnableBrowsing, iAbDefaultSearchBase, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_SERVICE_UID, PR_DISPLAY_NAME, PROP_AB_PROVIDER_SERVER_NAME, PROP_AB_PROVIDER_SERVER_PORT,
		PROP_AB_PROVIDER_USER_NAME, PROP_AB_PROVIDER_SEARCH_BASE, PROP_AB_PROVIDER_SEARCH_TIMEOUT, PROP_AB_PROVIDER_MAX_ENTRIES, PROP_AB_PROVIDER_USE_SSL,
		PROP_AB_PROVIDER_SERVER_SPA, PROP_AB_PROVIDER_ENABLE_BROWSING, PROP_AB_PROVIDER_SEARCH_BASE_DEFAULT };

	// Get access to the message service table, a list of the message services in the profile.
	CHK_HR_DBG(lpSvcAdmin2->GetMsgServiceTable(0, // Flags        
		&lpMsgSvcTable), L"lpSvcAdmin2->GetMsgServiceTable"); // Pointer to table

	// Set up restriction to query table.
	// Allocate and create the SRestriction
	// Allocate base memory:
	CHK_HR_DBG(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)& lpRes), L"MAPIAllocateBuffer");

	CHK_HR_DBG(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSvcName), L"MAPIAllocateMore");

	ZeroMemory(lpRes, sizeof(SRestriction));
	ZeroMemory(lpspvSvcName, sizeof(SPropValue));

	lpRes->rt = RES_CONTENT;
	lpRes->res.resContent.ulFuzzyLevel = FL_IGNORECASE | FL_FULLSTRING;
	lpRes->res.resContent.ulPropTag = PR_SERVICE_NAME_A;
	lpRes->res.resContent.lpProp = lpspvSvcName;
	lpspvSvcName->ulPropTag = PR_SERVICE_NAME_A;
	lpspvSvcName->Value.lpszA = ConvertWideCharToMultiByte(L"EMABLT");



	// Query the table to get the entry for EMABLT type services.
	CHK_HR_DBG(HrQueryAllRows(lpMsgSvcTable,
		(LPSPropTagArray)& sptaProps,
		lpRes,
		NULL,
		0,
		&lpSvcRows), L"HrQueryAllRows");


	if (lpSvcRows->cRows > 0)
	{
		for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
		{
			LPPROFSECT lpProfSect = NULL;
			LPMAPIPROP lpMapiProp = NULL;
			HCK(lpSvcAdmin2->OpenProfileSection(LPMAPIUID(lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb), NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSect));
			ULONG ulPropVal = 0;
			LPSPropValue lpsPropValues = NULL;
			HCK(lpProfSect->GetProps((LPSPropTagArray)& sptaProps, NULL, &ulPropVal, &lpsPropValues));
			if (lpsPropValues)
			{
				Logger::WriteLine(LOGLEVEL_INFO, L"  Listing entry #" + ConvertIntToString(i));
				Logger::WriteLine(LOGLEVEL_INFO, L"  Display Name        : " + ConvertWideCharToStdWstring(lpsPropValues[iDisplayName].Value.lpszW));
				Logger::WriteLine(LOGLEVEL_INFO, L"  Ldap Server Name    : " + ConvertMultiByteToStdWString(lpsPropValues[iAbServerName].Value.lpszA));
				Logger::WriteLine(LOGLEVEL_INFO, L"  Ldap Server Port    : " + ConvertMultiByteToStdWString(lpsPropValues[iAbServerPort].Value.lpszA));
				Logger::WriteLine(LOGLEVEL_INFO, L"  Username            : " + ConvertMultiByteToStdWString(lpsPropValues[iAbUsername].Value.lpszA));
				Logger::WriteLine(LOGLEVEL_INFO, L"  Search Base         : " + ConvertMultiByteToStdWString(lpsPropValues[iAbSearchBase].Value.lpszA));
				Logger::WriteLine(LOGLEVEL_INFO, L"  Search Timeout      : " + ConvertMultiByteToStdWString(lpsPropValues[iAbSearchTimeout].Value.lpszA));
				Logger::WriteLine(LOGLEVEL_INFO, L"  Maximum entries     : " + ConvertMultiByteToStdWString(lpsPropValues[iAbMaxEntries].Value.lpszA));
				if (lpsPropValues[iAbUseSSL].Value.b)
					Logger::WriteLine(LOGLEVEL_INFO, L"  Use SSL             : true");
				else
					Logger::WriteLine(LOGLEVEL_INFO, L"  Use SSL             : false");
				if (lpsPropValues[iAbRequireSpa].Value.b)
					Logger::WriteLine(LOGLEVEL_INFO, L"  Require SPA         : true");
				else
					Logger::WriteLine(LOGLEVEL_INFO, L"  Require SPA         : false");
				if (lpsPropValues[AbEnableBrowsing].Value.b)
					Logger::WriteLine(LOGLEVEL_INFO, L"  Enable browsing     : true");
				else
					Logger::WriteLine(LOGLEVEL_INFO, L"  Enable browsing     : false");
				if (lpsPropValues[iAbDefaultSearchBase].Value.i == 1)
					Logger::WriteLine(LOGLEVEL_INFO, L"  Default search base : true");
				else
					Logger::WriteLine(LOGLEVEL_INFO, L"  Default search base : false");
			}
			else
				Logger::WriteLine(LOGLEVEL_FAILED, L"Unable to retrieve Ldap AB properties.");
		}
	}
	else
		wprintf(L"No Ldap AB services found.\n");

Error:
	MAPIFreeBuffer(lpspvSvcName);
	MAPIFreeBuffer(lpRes);
	if (lpSvcRows) FreeProws(lpSvcRows);
	if (lpMsgSvcTable) lpMsgSvcTable->Release();
	return hRes;
}

HRESULT CreateService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID lpuidService)
{
	return S_OK;
}
// CreateABService
// Creates a new EMABLT service and populates the parameters
HRESULT CreateABService(LPSERVICEADMIN2 lpSvcAdmin2)
{
	HRESULT				hRes = S_OK;
	LPMAPITABLE			lpMsgSvcTable = NULL;		// MAPI table pointer.
	LPSRowSet			lpSvcRows = NULL;		// Row set pointer.
	SPropValue			rgval[12];						// Property value structure to hold configuration info.
	DATA_BLOB			dataBlobIn = { 0 };
	DATA_BLOB			dataBlobOut = { 0 };
	MAPIUID				uidService = { 0 };
	LPMAPIUID			lpuidService = &uidService;

	LPWSTR lpszwPassword = (LPWSTR)Toolkit::g_addressBookMap.at(L"password").c_str();
	LPTSTR lpszServiceName = (LPTSTR)Toolkit::g_addressBookMap.at(L"servicename").c_str();
	LPTSTR lpszDisplayName = (LPTSTR)Toolkit::g_addressBookMap.at(L"displayname").c_str();
	std::vector<SPropValue> rgvalVector;
	SPropValue sPropValue;

	// Adds a message service to the current profile and returns that newly added service UID.
	CHK_HR_DBG(lpSvcAdmin2->CreateMsgServiceEx((LPTSTR)ConvertWideCharToMultiByte((LPWSTR)Toolkit::g_addressBookMap.at(L"servicename").c_str()), (LPTSTR)ConvertWideCharToMultiByte((LPWSTR)Toolkit::g_addressBookMap.at(L"displayname").c_str()), NULL, 0, &uidService), L"Creating address book service" + Toolkit::g_addressBookMap.at(L"displayname"));

	rgvalVector.resize(0);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_DISPLAY_NAME;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"displayname").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_NAME;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"servername").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_PORT;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"serverport").c_str());
	rgvalVector.push_back(sPropValue);

	if (0 < wcslen(Toolkit::g_addressBookMap.at(L"username").c_str()))
	{
		ZeroMemory(&sPropValue, sizeof(SPropValue));
		sPropValue.ulPropTag = PROP_AB_PROVIDER_USER_NAME;
		sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"username").c_str());
		rgvalVector.push_back(sPropValue);
	}

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SEARCH_BASE;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"customsearchbase").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SEARCH_TIMEOUT;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"searchtimeout").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_MAX_ENTRIES;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"maxentries").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_USE_SSL;
	sPropValue.Value.b = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"usessl").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_SPA;
	sPropValue.Value.b = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"requirespa").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	// Encrypt the password if supplied
	if (0 < wcslen(lpszwPassword))
	{
		LPBYTE pbData = (LPBYTE)lpszwPassword;
		DWORD cbData = (wcslen(lpszwPassword) + 1) * sizeof(WCHAR);

		dataBlobIn.pbData = pbData;
		dataBlobIn.cbData = cbData;

		if (!CryptProtectData(
			&dataBlobIn,
			L"",						// desc
			NULL,						// optional
			NULL,						// reserver
			NULL,						// prompt struct
			0,							// flags
			&dataBlobOut))
		{
			wprintf(L"CryptProtectData failed!\n");
			hRes = E_FAIL;
			goto Error;
		}
		ZeroMemory(&sPropValue, sizeof(SPropValue));
		sPropValue.ulPropTag = PROP_AB_PROVIDER_USER_PASSWORD_ENCODED;
		sPropValue.Value.bin.cb = dataBlobOut.cbData;
		sPropValue.Value.bin.lpb = dataBlobOut.pbData;
		rgvalVector.push_back(sPropValue);
	}



	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_ENABLE_BROWSING;
	sPropValue.Value.b = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"enablebrowsing").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SEARCH_BASE_DEFAULT;
	sPropValue.Value.ul = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"defaultsearchbase").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	// Reconfigures a message service with the new props.
	CHK_HR_DBG(lpSvcAdmin2->ConfigureMsgService(lpuidService, NULL, 0, (ULONG)rgvalVector.size(), rgvalVector.data()), L"Configuring the address book service with the new properties");

Error:
	return hRes;
}

// GetABServiceUid
// Searches for an AB service with a given Display name and returns a service UID
HRESULT GetABServiceUid(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, ULONG * ulcMapiUid, MAPIUID* pMapiUid)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE		lpMsgSvcTable = NULL; // MAPI table pointer.
	LPSRowSet		lpSvcRows = NULL;
	LPSRestriction	lpRes = NULL;
	LPSRestriction	lpResLevel1 = NULL;
	LPSPropValue	lpspvSvcName = NULL;
	LPSPropValue	lpspvDispName = NULL;
	LPMAPIUID* pTempMAPIUid = NULL;
	enum { iServiceUid, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_SERVICE_UID };

	// Provides access to the message service table, a list of the message services in the profile.
	HCK(lpSvcAdmin2->GetMsgServiceTable(0, // Flags        
		&lpMsgSvcTable)); // Pointer to table

	// Set up restriction to query table.
	// Allocate and create our SRestriction
	// Allocate base memory:
	HCK(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)& lpRes));

	HCK(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpRes,
		(LPVOID*)& lpResLevel1));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSvcName));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvDispName));

	ZeroMemory(lpRes, sizeof(SRestriction));
	ZeroMemory(lpResLevel1, sizeof(SRestriction) * 2);

	ZeroMemory(lpspvSvcName, sizeof(SPropValue));
	ZeroMemory(lpspvDispName, sizeof(SPropValue));

	lpRes->rt = RES_AND;
	lpRes->res.resAnd.cRes = 2;
	lpRes->res.resAnd.lpRes = lpResLevel1;

	//Get the services matching the EMABLT service Name
	lpResLevel1[0].rt = RES_CONTENT;
	lpResLevel1[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
	lpResLevel1[0].res.resContent.ulPropTag = PR_SERVICE_NAME;
	lpResLevel1[0].res.resContent.lpProp = lpspvSvcName;
	lpspvSvcName->ulPropTag = PR_SERVICE_NAME;
	lpspvSvcName->Value.lpszA = ConvertWideCharToMultiByte(L"EMABLT");
	//Get the services matching the supplied display Name
	lpResLevel1[1].rt = RES_CONTENT;
	lpResLevel1[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
	lpResLevel1[1].res.resContent.ulPropTag = PR_DISPLAY_NAME;
	lpResLevel1[1].res.resContent.lpProp = lpspvDispName;
	lpspvDispName->ulPropTag = PR_DISPLAY_NAME;
	lpspvDispName->Value.lpszA = ConvertWideCharToMultiByte(lppszDisplayName);

	// Query the table to get the entry for the EMABLT service.
	CHK_HR_DBG(HrQueryAllRows(lpMsgSvcTable,
		(LPSPropTagArray)& sptaProps,
		lpRes,
		NULL,
		0,
		&lpSvcRows), L"Querrying service rows");

	if (lpSvcRows->cRows > 0)
	{
		if (ulcMapiUid)
			* ulcMapiUid = lpSvcRows->cRows;
		if (pMapiUid)
			for (int i = 0; i < lpSvcRows->cRows; i++)
			{
				*(&pMapiUid[i]) = *(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb;
			}
	}

Error:
	MAPIFreeBuffer(lpspvDispName);
	MAPIFreeBuffer(lpspvSvcName);
	MAPIFreeBuffer(lpResLevel1);
	MAPIFreeBuffer(lpRes);
	if (lpSvcRows) FreeProws(lpSvcRows);
	if (lpMsgSvcTable) lpMsgSvcTable->Release();
	return hRes;
}

// GetABServiceUid
// Searches for an AB service with a given Display name and Server name and returns a service UID
HRESULT GetABServiceUid(LPSERVICEADMIN2 lpSvcAdmin2, ULONG* ulcMapiUid, MAPIUID* pMapiUid)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE		lpMsgSvcTable = NULL; // MAPI table pointer.
	LPSRowSet		lpSvcRows = NULL;
	LPSRestriction	lpRes = NULL;
	LPMAPIUID* pTempMAPIUid = NULL;
	LPSPropValue	lpspvSvcName = NULL;
	enum { iServiceUid, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_SERVICE_UID };

	// Provides access to the message service table, a list of the message services in the profile.
	HCK(lpSvcAdmin2->GetMsgServiceTable(0, // Flags        
		&lpMsgSvcTable)); // Pointer to table

	// Set up restriction to query table.

	// Allocate and create our SRestriction
	// Allocate base memory:
	HCK(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)& lpRes));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSvcName));

	ZeroMemory(lpRes, sizeof(SRestriction));

	lpRes->rt = RES_CONTENT;
	lpRes->res.resContent.ulFuzzyLevel = FL_FULLSTRING;
	lpRes->res.resContent.ulPropTag = PR_SERVICE_NAME_A;
	lpRes->res.resContent.lpProp = lpspvSvcName;
	lpspvSvcName->ulPropTag = PR_SERVICE_NAME_A;
	lpspvSvcName->Value.lpszA = ConvertWideCharToMultiByte(L"EMABLT");

	// Query the table to get the entry for the Exchange message service.
	CHK_HR_DBG(HrQueryAllRows(lpMsgSvcTable,
		(LPSPropTagArray)& sptaProps,
		lpRes,
		NULL,
		0,
		&lpSvcRows), L"Querying service rows");

	if (lpSvcRows->cRows > 0)
	{
		if (ulcMapiUid)
			* ulcMapiUid = lpSvcRows->cRows;
		if (pMapiUid)
			for (int i = 0; i < lpSvcRows->cRows; i++)
			{
				*(&pMapiUid[i]) = *(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb;
			}
	}

Error:
	MAPIFreeBuffer(lpspvSvcName);
	MAPIFreeBuffer(lpRes);
	if (lpSvcRows) FreeProws(lpSvcRows);
	if (lpMsgSvcTable) lpMsgSvcTable->Release();
	return hRes;
}


// GetABServiceUid
// Searches for an AB service with a given Display name and Server name and returns a service UID
HRESULT GetABServiceUid(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, LPTSTR lppszServerName, ULONG* ulcMapiUid, MAPIUID* pMapiUid)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE		lpMsgSvcTable = NULL; // MAPI table pointer.
	LPSRowSet		lpSvcRows = NULL;
	LPSRestriction	lpRes = NULL;
	LPSRestriction	lpResLevel0 = NULL;
	LPSRestriction	lpResLevel1 = NULL;
	LPSPropValue	lpspvSvcName = NULL;
	LPSPropValue	lpspvDispName = NULL;
	LPSPropValue	lpspvSrvName = NULL;

	enum { iServiceUid, iDisplayName, iServerName, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_SERVICE_UID, PROP_AB_PROVIDER_DISPLAY_NAME, PROP_AB_PROVIDER_SERVER_NAME };
	// Provides access to the message service table, a list of the message services in the profile.
	hRes = lpSvcAdmin2->GetMsgServiceTable(0, // Flags        
		&lpMsgSvcTable); // Pointer to table
	if (FAILED(hRes)) goto Error;

	// Set up restriction to query table.

	// Allocate and create our SRestriction
	// Allocate base memory:
	HCK(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)& lpRes));

	HCK(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpRes,
		(LPVOID*)& lpResLevel0));

	HCK(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpRes,
		(LPVOID*)& lpResLevel1));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSvcName));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvDispName));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSrvName));

	ZeroMemory(lpRes, sizeof(SRestriction));
	ZeroMemory(lpResLevel0, sizeof(SRestriction) * 2);
	ZeroMemory(lpResLevel1, sizeof(SRestriction) * 2);
	ZeroMemory(lpspvSvcName, sizeof(SPropValue));
	ZeroMemory(lpspvDispName, sizeof(SPropValue));
	ZeroMemory(lpspvSrvName, sizeof(SPropValue));

	lpRes->rt = RES_AND;
	lpRes->res.resAnd.cRes = 2;
	lpRes->res.resAnd.lpRes = lpResLevel0;
	//Get the services matching the EMABLT Service Name

	lpResLevel0[0].rt = RES_EXIST;
	lpResLevel0[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;

	lpResLevel0[1].rt = RES_CONTENT;
	lpResLevel0[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
	lpResLevel0[1].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
	lpResLevel0[1].res.resContent.lpProp = lpspvSvcName;
	lpspvSvcName->ulPropTag = PR_SERVICE_NAME_A;
	lpspvSvcName->Value.lpszA = ConvertWideCharToMultiByte(L"EMABLT");




	// Query the table to get the entry for the Exchange message service.
	HCK(HrQueryAllRows(lpMsgSvcTable,
		(LPSPropTagArray)& sptaProps,
		lpRes,
		NULL,
		0,
		&lpSvcRows));

	if (lpSvcRows->cRows > 0)
	{
		// this is a not so pretty workaround for my not being able to get the MAPI table restriction to work with PROP_AB_PROVIDER_SERVER_NAME
		// MFCMAPI also returns 0 rows which is strange
		ULONG cEntries = 0;



		if (ulcMapiUid)
		{
			for (int i = 0; i < lpSvcRows->cRows; i++)
			{
				LPPROFSECT lpProfSect = NULL;
				LPMAPIPROP lpMapiProp = NULL;
				HCK(lpSvcAdmin2->OpenProfileSection(LPMAPIUID(lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb), NULL, MAPI_FORCE_ACCESS, &lpProfSect));
				ULONG ulPropVal = 0;
				LPSPropValue lpsPropValues = NULL;
				HCK(lpProfSect->GetProps((LPSPropTagArray)& sptaProps, NULL, &ulPropVal, &lpsPropValues));
				if (lppszDisplayName && lppszServerName)
				{
					if (0 == _stricmp(ConvertWideCharToMultiByte(lppszDisplayName), lpsPropValues[iDisplayName].Value.lpszA))
					{
						cEntries++;
					}
					else if (0 == _stricmp(ConvertWideCharToMultiByte(lppszServerName), lpsPropValues[iServerName].Value.lpszA))
					{
						cEntries++;
					}
				}
				else if (lppszDisplayName)
				{
					if (0 == _stricmp(ConvertWideCharToMultiByte(lppszDisplayName), lpsPropValues[iDisplayName].Value.lpszA))
					{
						cEntries++;
					}
				}
				else if (lppszServerName)
				{
					if (0 == _stricmp(ConvertWideCharToMultiByte(lppszServerName), lpsPropValues[iServerName].Value.lpszA))
					{
						cEntries++;
					}
				}
				if (lpsPropValues) MAPIFreeBuffer(lpsPropValues);
				if (lpProfSect) lpProfSect->Release();
			}

			*ulcMapiUid = cEntries;
		}
		if (pMapiUid)
			for (int i = 0; i < lpSvcRows->cRows; i++)
			{
				LPPROFSECT lpProfSect = NULL;
				LPMAPIPROP lpMapiProp = NULL;
				HCK(lpSvcAdmin2->OpenProfileSection(LPMAPIUID(lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb), NULL, MAPI_FORCE_ACCESS, &lpProfSect));
				ULONG ulPropVal = 0;
				LPSPropValue lpsPropValues = NULL;
				HCK(lpProfSect->GetProps((LPSPropTagArray)& sptaProps, NULL, &ulPropVal, &lpsPropValues));
				if (lppszDisplayName && lppszServerName)
				{
					if (0 == _stricmp(ConvertWideCharToMultiByte(lppszDisplayName), lpsPropValues[iDisplayName].Value.lpszA))
					{
						*(&pMapiUid[cEntries]) = *(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb;
						cEntries++;
					}
					else if (0 == _stricmp(ConvertWideCharToMultiByte(lppszServerName), lpsPropValues[iServerName].Value.lpszA))
					{
						*(&pMapiUid[cEntries]) = *(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb;
						cEntries++;
					}
				}
				else if (lppszDisplayName)
				{
					if (0 == _stricmp(ConvertWideCharToMultiByte(lppszDisplayName), lpsPropValues[iDisplayName].Value.lpszA))
					{
						*(&pMapiUid[cEntries]) = *(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb;
						cEntries++;
					}
				}
				else if (lppszServerName)
				{
					if (0 == _stricmp(ConvertWideCharToMultiByte(lppszServerName), lpsPropValues[iServerName].Value.lpszA))
					{
						*(&pMapiUid[cEntries]) = *(LPMAPIUID)lpSvcRows->aRow[i].lpProps[iServiceUid].Value.bin.lpb;
						cEntries++;
					}
				}

				if (lpsPropValues) MAPIFreeBuffer(lpsPropValues);
				if (lpProfSect) lpProfSect->Release();
			}


	}


Error:
	if (lpspvSrvName) MAPIFreeBuffer(lpspvSrvName);
	if (lpspvDispName) MAPIFreeBuffer(lpspvDispName);
	if (lpspvSvcName) MAPIFreeBuffer(lpspvSvcName);
	if (lpResLevel1) MAPIFreeBuffer(lpResLevel1);
	if (lpRes) MAPIFreeBuffer(lpRes);
	if (lpSvcRows) FreeProws(lpSvcRows);
	if (lpMsgSvcTable) lpMsgSvcTable->Release();
	return hRes;
}

// UpdateABService
// Updates the AB service with the given service UID
HRESULT UpdateABService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID lpMapiUid)
{
	HRESULT				hRes = S_OK;
	LPMAPITABLE			lpMsgSvcTable = NULL;		// MAPI table pointer.
	LPSRowSet			lpSvcRows = NULL;		// Row set pointer.
	SPropValue			rgval[12];						// Property value structure to hold configuration info.
	DATA_BLOB			dataBlobIn = { 0 };
	DATA_BLOB			dataBlobOut = { 0 };

	LPWSTR lpszwPassword = (LPWSTR)Toolkit::g_addressBookMap.at(L"password").c_str();
	std::vector<SPropValue> rgvalVector;
	SPropValue sPropValue;

	rgvalVector.resize(0);

	if (!Toolkit::g_addressBookMap.at(L"newdisplayname").empty())
	{
	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_DISPLAY_NAME;
		sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"newdisplayname").c_str());
		rgvalVector.push_back(sPropValue);
		Toolkit::g_addressBookMap.at(L"displayname") = Toolkit::g_addressBookMap.at(L"newdisplayname");
	}
	else
	{
		if (!Toolkit::g_addressBookMap.at(L"displayname").empty())
		{
			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PROP_AB_PROVIDER_DISPLAY_NAME;
			sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"displayname").c_str());
			rgvalVector.push_back(sPropValue);
		}
	}

	if (!Toolkit::g_addressBookMap.at(L"newservername").empty())
	{
	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_NAME;
		sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"newservername").c_str());
		rgvalVector.push_back(sPropValue);
		Toolkit::g_addressBookMap.at(L"servername") = Toolkit::g_addressBookMap.at(L"newservername");
	}
	else
	{
		if (!Toolkit::g_addressBookMap.at(L"servername").empty())
		{
			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_NAME;
			sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"servername").c_str());
			rgvalVector.push_back(sPropValue);
		}

	}
	
	if (!Toolkit::g_addressBookMap.at(L"newserverport").empty())
	{
	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_PORT;
		sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"newserverport").c_str());
		rgvalVector.push_back(sPropValue);
		Toolkit::g_addressBookMap.at(L"serverport") = Toolkit::g_addressBookMap.at(L"newserverport");
	}
	else
	{
		if (!Toolkit::g_addressBookMap.at(L"serverport").empty())
		{
			ZeroMemory(&sPropValue, sizeof(SPropValue));
			sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_PORT;
			sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"serverport").c_str());
			rgvalVector.push_back(sPropValue);
		}
	}

	if (0 < wcslen(Toolkit::g_addressBookMap.at(L"username").c_str()))
	{
		ZeroMemory(&sPropValue, sizeof(SPropValue));
		sPropValue.ulPropTag = PROP_AB_PROVIDER_USER_NAME;
		sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"username").c_str());
		rgvalVector.push_back(sPropValue);
	}

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SEARCH_BASE;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"customsearchbase").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SEARCH_TIMEOUT;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"searchtimeout").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_MAX_ENTRIES;
	sPropValue.Value.lpszA = ConvertWideCharToMultiByte(Toolkit::g_addressBookMap.at(L"maxentries").c_str());
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_USE_SSL;
	sPropValue.Value.b = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"usessl").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SERVER_SPA;
	sPropValue.Value.b = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"requirespa").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	// Encrypt the password if supplied
	if (0 < wcslen(lpszwPassword))
	{
		LPBYTE pbData = (LPBYTE)lpszwPassword;
		DWORD cbData = (wcslen(lpszwPassword) + 1) * sizeof(WCHAR);

		dataBlobIn.pbData = pbData;
		dataBlobIn.cbData = cbData;

		if (!CryptProtectData(
			&dataBlobIn,
			L"",						// desc
			NULL,						// optional
			NULL,						// reserver
			NULL,						// prompt struct
			0,							// flags
			&dataBlobOut))
		{
			wprintf(L"CryptProtectData failed!\n");
			hRes = E_FAIL;
			goto Error;
		}
		ZeroMemory(&sPropValue, sizeof(SPropValue));
		sPropValue.ulPropTag = PROP_AB_PROVIDER_USER_PASSWORD_ENCODED;
		sPropValue.Value.bin.cb = dataBlobOut.cbData;
		sPropValue.Value.bin.lpb = dataBlobOut.pbData;
		rgvalVector.push_back(sPropValue);
	}



	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_ENABLE_BROWSING;
	sPropValue.Value.b = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"enablebrowsing").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	ZeroMemory(&sPropValue, sizeof(SPropValue));
	sPropValue.ulPropTag = PROP_AB_PROVIDER_SEARCH_BASE_DEFAULT;
	sPropValue.Value.ul = (0 == wcscmp(Toolkit::g_addressBookMap.at(L"defaultsearchbase").c_str(), L"true"));
	rgvalVector.push_back(sPropValue);

	// Reconfigures a message service with the new props.
	CHK_HR_DBG(lpSvcAdmin2->ConfigureMsgService(lpMapiUid, NULL, 0, (ULONG)rgvalVector.size(), rgvalVector.data()), L"ConfigureMsgService");

Error:
	return hRes;
}

// RemoveABService
// Removes the AB sercie with the given service UID
HRESULT RemoveABService(LPSERVICEADMIN2 lpSvcAdmin2, LPMAPIUID lpMapiUid)
{
	HRESULT hRes = S_OK;
	// Deletes a message service from a profile.
	CHK_HR_DBG(lpSvcAdmin2->DeleteMsgService(lpMapiUid), L"DeleteMsgService");
Error:
	
	return hRes;
}

HRESULT CheckABServiceExists(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, LPTSTR lppszServerName, BOOL* success)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE		lpMsgSvcTable = NULL; // MAPI table pointer.
	LPSRowSet		lpSvcRows = NULL;
	LPSRestriction	lpRes = NULL;
	LPSRestriction	lpResLevel0 = NULL;
	LPSRestriction	lpResLevel1 = NULL;
	LPSPropValue	lpspvSvcName = NULL;
	LPSPropValue	lpspvDispName = NULL;
	LPSPropValue	lpspvSrvName = NULL;
	LPMAPIUID* pTempMAPIUid = NULL;
	enum { iServiceUid, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_SERVICE_UID };
	
	// Provides access to the message service table, a list of the message services in the profile.
	hRes = lpSvcAdmin2->GetMsgServiceTable(0, // Flags        
		&lpMsgSvcTable); // Pointer to table
	if (FAILED(hRes)) goto Error;

	// Set up restriction to query table.

	// Allocate and create our SRestriction
	// Allocate base memory:
	HCK(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)& lpRes));

	HCK(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpRes,
		(LPVOID*)& lpResLevel0));

	HCK(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpRes,
		(LPVOID*)& lpResLevel1));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSvcName));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvDispName));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSrvName));

	ZeroMemory(lpRes, sizeof(SRestriction));
	ZeroMemory(lpResLevel0, sizeof(SRestriction) * 2);
	ZeroMemory(lpResLevel1, sizeof(SRestriction) * 2);
	ZeroMemory(lpspvSvcName, sizeof(SPropValue));
	ZeroMemory(lpspvDispName, sizeof(SPropValue));
	ZeroMemory(lpspvSrvName, sizeof(SPropValue));

	lpRes->rt = RES_AND;
	lpRes->res.resAnd.cRes = 2;
	lpRes->res.resAnd.lpRes = lpResLevel0;
	//Get the services matching the EMABLT Service Name

	lpResLevel0[0].rt = RES_CONTENT;
	lpResLevel0[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
	lpResLevel0[0].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
	lpResLevel0[0].res.resContent.lpProp = lpspvSvcName;
	lpspvSvcName->ulPropTag = PR_SERVICE_NAME_A;
	lpspvSvcName->Value.lpszA = ConvertWideCharToMultiByte(L"EMABLT");
	//Get the services matching the supplied Display Name

	lpResLevel0[1].rt = RES_OR;
	lpResLevel0[1].res.resOr.cRes = 2;
	lpResLevel0[1].res.resOr.lpRes = lpResLevel1;

	lpResLevel1[0].rt = RES_CONTENT;
	lpResLevel1[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING | FL_IGNORECASE;
	lpResLevel1[0].res.resContent.ulPropTag = PR_DISPLAY_NAME_A;
	lpResLevel1[0].res.resContent.lpProp = lpspvDispName;
	lpspvDispName->ulPropTag = PR_DISPLAY_NAME_A;
	lpspvDispName->Value.lpszA = ConvertWideCharToMultiByte(lppszDisplayName);
	//Get the services matching the supplied ldap server name
	lpResLevel1[1].rt = RES_CONTENT;
	lpResLevel1[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING | FL_IGNORECASE;
	lpResLevel1[1].res.resContent.ulPropTag = PROP_AB_PROVIDER_SERVER_NAME;
	lpResLevel1[1].res.resContent.lpProp = lpspvSrvName;
	lpspvSrvName->ulPropTag = PROP_AB_PROVIDER_SERVER_NAME;
	lpspvSrvName->Value.lpszA = ConvertWideCharToMultiByte(lppszServerName);

	// Query the table to get the entry for the EMABLT service.
	HCK(HrQueryAllRows(lpMsgSvcTable,
		(LPSPropTagArray)& sptaProps,
		lpRes,
		NULL,
		0,
		&lpSvcRows));

	if (lpSvcRows->cRows > 0)
	{
		*success = true;
	}
	else
		*success = false;

Error:
	MAPIFreeBuffer(lpspvDispName);
	MAPIFreeBuffer(lpspvSvcName);
	MAPIFreeBuffer(lpResLevel1);
	MAPIFreeBuffer(lpRes);
	if (lpSvcRows) FreeProws(lpSvcRows);
	if (lpMsgSvcTable) lpMsgSvcTable->Release();
	return hRes;
}
HRESULT CheckABServiceExists(LPSERVICEADMIN2 lpSvcAdmin2, LPTSTR lppszDisplayName, BOOL* success)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE		lpMsgSvcTable = NULL; // MAPI table pointer.
	LPSRowSet		lpSvcRows = NULL;
	LPSRestriction	lpRes = NULL;
	LPSRestriction	lpResLevel1 = NULL;
	LPSPropValue	lpspvSvcName = NULL;
	LPSPropValue	lpspvDispName = NULL;

	enum { iServiceUid, cptaProps };
	SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_SERVICE_UID };

	// Provides access to the message service table, a list of the message services in the profile.
	HCK(lpSvcAdmin2->GetMsgServiceTable(0, // Flags        
		&lpMsgSvcTable)); // Pointer to table

	// Set up restriction to query table.
	// Allocate and create our SRestriction
	// Allocate base memory:
	HCK(MAPIAllocateBuffer(
		sizeof(SRestriction),
		(LPVOID*)& lpRes));

	HCK(MAPIAllocateMore(
		sizeof(SRestriction) * 2,
		lpRes,
		(LPVOID*)& lpResLevel1));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvSvcName));

	HCK(MAPIAllocateMore(
		sizeof(SPropValue),
		lpRes,
		(LPVOID*)& lpspvDispName));

	ZeroMemory(lpRes, sizeof(SRestriction));
	ZeroMemory(lpResLevel1, sizeof(SRestriction) * 2);

	ZeroMemory(lpspvSvcName, sizeof(SPropValue));
	ZeroMemory(lpspvDispName, sizeof(SPropValue));

	lpRes->rt = RES_AND;
	lpRes->res.resAnd.cRes = 2;
	lpRes->res.resAnd.lpRes = lpResLevel1;

	//Get the services matching the EMABLT service Name
	lpResLevel1[0].rt = RES_CONTENT;
	lpResLevel1[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
	lpResLevel1[0].res.resContent.ulPropTag = PR_SERVICE_NAME;
	lpResLevel1[0].res.resContent.lpProp = lpspvSvcName;
	lpspvSvcName->ulPropTag = PR_SERVICE_NAME;
	lpspvSvcName->Value.lpszA = ConvertWideCharToMultiByte(L"EMABLT");
	//Get the services matching the supplied display Name
	lpResLevel1[1].rt = RES_CONTENT;
	lpResLevel1[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
	lpResLevel1[1].res.resContent.ulPropTag = PR_DISPLAY_NAME;
	lpResLevel1[1].res.resContent.lpProp = lpspvDispName;
	lpspvDispName->ulPropTag = PR_DISPLAY_NAME;
	lpspvDispName->Value.lpszA = ConvertWideCharToMultiByte(lppszDisplayName);

	// Query the table to get the entry for the EMABLT service.
	HCK(HrQueryAllRows(lpMsgSvcTable,
		(LPSPropTagArray)& sptaProps,
		lpRes,
		NULL,
		0,
		&lpSvcRows));

	if (lpSvcRows->cRows > 0)
	{
		*success = true;
	}
	else
		*success = false;

Error:
	MAPIFreeBuffer(lpspvDispName);
	MAPIFreeBuffer(lpspvSvcName);
	MAPIFreeBuffer(lpResLevel1);
	MAPIFreeBuffer(lpRes);
	if (lpSvcRows) FreeProws(lpSvcRows);
	if (lpMsgSvcTable) lpMsgSvcTable->Release();
	return hRes;
}
}