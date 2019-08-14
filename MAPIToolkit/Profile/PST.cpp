#include "PST.h"
#include "../InlineAndMacros.h"
#include <MAPIUtil.h>
#include "../ExtraMAPIDefs.h"
#include "..//Misc/Utility/StringOperations.h"
#include "shlwapi.h"

namespace MAPIToolkit
{
	HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszOldPath, LPWSTR lpszNewPath, bool bMoveFiles)
	{
		HRESULT hRes = S_OK;

		LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
		LPSRestriction lpProfRes = NULL;
		LPSRestriction lpProfResLvl1 = NULL;
		LPSPropValue lpProfPropVal = NULL;
		LPMAPITABLE lpProfTable = NULL;
		LPSRowSet lpProfRows = NULL;
		LPSERVICEADMIN lpServiceAdmin = NULL;
		LPMAPITABLE lpServiceTable = NULL;
		// Setting up an enum and a prop tag array with the props we'll use
		enum { iDisplayName, cptaProps };
		SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME };

		CHK_HR_DBG(MAPIAdminProfiles(0, // Flags
			&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
										 // Get an IProfAdmin interface.

		CHK_HR_DBG(lpProfAdmin->GetProfileTable(0,
			&lpProfTable), L"Calling GetProfileTable");

		// Allocate memory for the restriction
		CHK_HR_DBG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)& lpProfRes), L"Calling MAPIAllocateBuffer");

		CHK_HR_DBG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)& lpProfResLvl1), L"Calling MAPIAllocateBuffer");

		CHK_HR_DBG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)& lpProfPropVal), L"Calling MAPIAllocateBuffer");

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
			(LPSPropTagArray)& sptaProps,
			lpProfRes,
			NULL,
			0,
			&lpProfRows), L"Calling #");

		if (lpProfRows->cRows == 0)
		{
			return MAPI_E_NOT_FOUND;
		}
		else if (lpProfRows->cRows != 1)
		{
			return MAPI_E_CALL_FAILED;
		}

		// Begin process services

		CHK_HR_DBG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
			LPTSTR(""),            // Password for that profile.
			NULL,                // Handle to parent window.
			MAPI_UNICODE,                    // Flags.
			&lpServiceAdmin), L"Calling #");        // Pointer to new IMsgServiceAdmin.

		if (lpServiceAdmin)
		{
			lpServiceAdmin->GetMsgServiceTable(0,
				&lpServiceTable);
			LPSRestriction lpSvcRes = NULL;
			LPSRestriction lpsvcResLvl1 = NULL;
			LPSRestriction lpsvcResLvl2 = NULL;
			LPSPropValue lpSvcPropVal1 = NULL;
			LPSPropValue lpSvcPropVal2 = NULL;
			LPSRowSet lpSvcRows = NULL;

			// Setting up an enum and a prop tag array with the props we'll use
			enum { iServiceUid, iServiceName, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
			SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID,PR_SERVICE_NAME_A, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

			// Allocate memory for the restriction
			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SRestriction),
				(LPVOID*)& lpSvcRes), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SRestriction) * 2,
				(LPVOID*)& lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SRestriction) * 2,
				(LPVOID*)& lpsvcResLvl2), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SPropValue),
				(LPVOID*)& lpSvcPropVal1), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SPropValue),
				(LPVOID*)& lpSvcPropVal2), L"Calling MAPIAllocateBuffer");

			// Set up restriction to query the profile table
			lpSvcRes->rt = RES_AND;
			lpSvcRes->res.resAnd.cRes = 0x00000002;
			lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

			lpsvcResLvl1[0].rt = RES_EXIST;
			lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
			lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
			lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;

			lpsvcResLvl1[1].rt = RES_OR;
			lpsvcResLvl1[1].res.resOr.cRes = 0x00000002;
			lpsvcResLvl1[1].res.resOr.lpRes = lpsvcResLvl2;

			lpsvcResLvl2[0].rt = RES_CONTENT;
			lpsvcResLvl2[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
			lpsvcResLvl2[0].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
			lpsvcResLvl2[0].res.resContent.lpProp = lpSvcPropVal1;

			lpSvcPropVal1->ulPropTag = PR_SERVICE_NAME_A;
			lpSvcPropVal1->Value.lpszA = ConvertWideCharToMultiByte(L"MSPST MS");

			lpsvcResLvl2[1].rt = RES_CONTENT;
			lpsvcResLvl2[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
			lpsvcResLvl2[1].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
			lpsvcResLvl2[1].res.resContent.lpProp = lpSvcPropVal2;

			lpSvcPropVal2->ulPropTag = PR_SERVICE_NAME_A;
			lpSvcPropVal2->Value.lpszA = ConvertWideCharToMultiByte(L"MSUPST MS");

			// Query the table to get the the default profile only
			CHK_HR_DBG(HrQueryAllRows(lpServiceTable,
				(LPSPropTagArray)& sptaSvcProps,
				lpSvcRes,
				NULL,
				0,
				&lpSvcRows), L"Calling HrQueryAllRows");

			if (lpSvcRows->cRows > 0)
			{
				wprintf(L"Found %i PST services in profile %s\n", lpSvcRows->cRows, lpszProfileName);
				// Start loop services
				for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
				{

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
							(LPVOID*)& lpProvRes), L"Calling MAPIAllocateBuffer");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)& lpProvResLvl1), L"Calling MAPIAllocateBuffer");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)& lpProvPropVal), L"Calling MAPIAllocateBuffer");

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
							(LPSPropTagArray)& sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), L"Calling HrQueryAllRows");

						if (lpProvRows->cRows > 0)
						{

							LPPROFSECT lpProfSection = NULL;
							if (SUCCEEDED(lpServiceAdmin->OpenProfileSection((LPMAPIUID)lpProvRows->aRow->lpProps[iProvInstanceKey].Value.bin.lpb, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSection)))
							{
								LPMAPIPROP lpMAPIProp = NULL;
								if (SUCCEEDED(lpProfSection->QueryInterface(IID_IMAPIProp, (void**)& lpMAPIProp)))
								{
									LPSPropValue prDisplayName = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_DISPLAY_NAME_W, &prDisplayName)))
									{
										// bind to the PR_PST_PATH_W property
										LPSPropValue pstPathW = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_PATH_W, &pstPathW)))
										{
											if (pstPathW)
											{
												std::wstring szCurrentPath = ConvertWideCharToStdWstring(pstPathW->Value.lpszW);
												if (WStringReplace(&szCurrentPath, ConvertWideCharToStdWstring(lpszOldPath), ConvertWideCharToStdWstring(lpszNewPath)))
												{
													if (bMoveFiles)
													{
														wprintf(L"Moving file %s to new location %s\n", pstPathW->Value.lpszW, szCurrentPath.c_str());
														BOOL bFileMoved = false;
														bFileMoved = MoveFileExW(pstPathW->Value.lpszW, (LPCWSTR)szCurrentPath.c_str(), MOVEFILE_COPY_ALLOWED | MOVEFILE_WRITE_THROUGH);
														if (bFileMoved)
														{
															wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
															pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
															CHK_HR_DBG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
														}
														else
														{
															wprintf(L"Unable to move file\n");
														}
													}
													else
													{
														wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
														pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
														CHK_HR_DBG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
													}
												}
												if (pstPathW) MAPIFreeBuffer(pstPathW);
											}
										}
										if (prDisplayName) MAPIFreeBuffer(prDisplayName);
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
				if (lpSvcRows) FreeProws(lpSvcRows);
				// End loop services
			}
			if (lpSvcPropVal1) MAPIFreeBuffer(lpSvcPropVal1);
			if (lpSvcPropVal2) MAPIFreeBuffer(lpSvcPropVal2);
			if (lpsvcResLvl1) MAPIFreeBuffer(lpsvcResLvl1);
			if (lpsvcResLvl2) MAPIFreeBuffer(lpsvcResLvl2);
			if (lpSvcRes) MAPIFreeBuffer(lpSvcRes);
			if (lpServiceTable) lpServiceTable->Release();
			if (lpServiceAdmin) lpServiceAdmin->Release();
		}
		// End process services

		// Free up memory
		//if (lpProfPropVal) MAPIFreeBuffer(lpProfPropVal);
		//if (lpProfResLvl1) MAPIFreeBuffer(lpProfResLvl1);
		//if (lpProfRes) MAPIFreeBuffer(lpProfRes);
		if (lpProfRows) FreeProws(lpProfRows);
		if (lpProfTable) lpProfTable->Release();
		if (lpProfAdmin) lpProfAdmin->Release();
	Error:
		goto CleanUp;

	CleanUp:
		return hRes;
	}

	HRESULT UpdatePstPath(LPWSTR lpszProfileName, LPWSTR lpszNewPath, bool bMoveFiles)
	{
		HRESULT hRes = S_OK;

		LPPROFADMIN lpProfAdmin = NULL;     // Profile Admin pointer
		LPSRestriction lpProfRes = NULL;
		LPSRestriction lpProfResLvl1 = NULL;
		LPSPropValue lpProfPropVal = NULL;
		LPMAPITABLE lpProfTable = NULL;
		LPSRowSet lpProfRows = NULL;

		LPSERVICEADMIN lpServiceAdmin = NULL;
		LPMAPITABLE lpServiceTable = NULL;
		// Setting up an enum and a prop tag array with the props we'll use
		enum { iDisplayName, cptaProps };
		SizedSPropTagArray(cptaProps, sptaProps) = { cptaProps, PR_DISPLAY_NAME };

		CHK_HR_DBG(MAPIAdminProfiles(0, // Flags
			&lpProfAdmin), L"Calling MAPIAdminProfiles"); // Pointer to new IProfAdmin
										 // Get an IProfAdmin interface.

		CHK_HR_DBG(lpProfAdmin->GetProfileTable(0,
			&lpProfTable), L"Calling GetProfileTable");

		// Allocate memory for the restriction
		CHK_HR_DBG(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)& lpProfRes), L"Calling MAPIAllocateBuffer");

		CHK_HR_DBG(MAPIAllocateBuffer(
			sizeof(SRestriction) * 2,
			(LPVOID*)& lpProfResLvl1), L"Calling MAPIAllocateBuffer");

		CHK_HR_DBG(MAPIAllocateBuffer(
			sizeof(SPropValue),
			(LPVOID*)& lpProfPropVal), L"Calling MAPIAllocateBuffer");

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
			(LPSPropTagArray)& sptaProps,
			lpProfRes,
			NULL,
			0,
			&lpProfRows), L"Calling HrQueryAllRows");

		if (lpProfRows->cRows == 0)
		{
			return MAPI_E_NOT_FOUND;
		}
		else if (lpProfRows->cRows != 1)
		{
			return MAPI_E_CALL_FAILED;
		}

		// Begin process services

		CHK_HR_DBG(lpProfAdmin->AdminServices((LPTSTR)lpszProfileName,
			LPTSTR(""),            // Password for that profile.
			NULL,                // Handle to parent window.
			MAPI_UNICODE,                    // Flags.
			&lpServiceAdmin), L"Calling AdminServices");        // Pointer to new IMsgServiceAdmin.

		if (lpServiceAdmin)
		{
			lpServiceAdmin->GetMsgServiceTable(0,
				&lpServiceTable);
			LPSRestriction lpSvcRes = NULL;
			LPSRestriction lpsvcResLvl1 = NULL;
			LPSRestriction lpsvcResLvl2 = NULL;
			LPSPropValue lpSvcPropVal1 = NULL;
			LPSPropValue lpSvcPropVal2 = NULL;
			LPSRowSet lpSvcRows = NULL;

			// Setting up an enum and a prop tag array with the props we'll use
			enum { iServiceUid, iServiceName, iEmsMdbSectUid, iServiceResFlags, cptaSvcProps };
			SizedSPropTagArray(cptaSvcProps, sptaSvcProps) = { cptaSvcProps, PR_SERVICE_UID,PR_SERVICE_NAME_A, PR_EMSMDB_SECTION_UID, PR_RESOURCE_FLAGS };

			// Allocate memory for the restriction
			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SRestriction),
				(LPVOID*)& lpSvcRes), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SRestriction) * 2,
				(LPVOID*)& lpsvcResLvl1), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SRestriction) * 2,
				(LPVOID*)& lpsvcResLvl2), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SPropValue),
				(LPVOID*)& lpSvcPropVal1), L"Calling MAPIAllocateBuffer");

			CHK_HR_DBG(MAPIAllocateBuffer(
				sizeof(SPropValue),
				(LPVOID*)& lpSvcPropVal2), L"Calling MAPIAllocateBuffer");

			// Set up restriction to query the profile table
			lpSvcRes->rt = RES_AND;
			lpSvcRes->res.resAnd.cRes = 0x00000002;
			lpSvcRes->res.resAnd.lpRes = lpsvcResLvl1;

			lpsvcResLvl1[0].rt = RES_EXIST;
			lpsvcResLvl1[0].res.resExist.ulPropTag = PR_SERVICE_NAME_A;
			lpsvcResLvl1[0].res.resExist.ulReserved1 = 0x00000000;
			lpsvcResLvl1[0].res.resExist.ulReserved2 = 0x00000000;

			lpsvcResLvl1[1].rt = RES_OR;
			lpsvcResLvl1[1].res.resOr.cRes = 0x00000002;
			lpsvcResLvl1[1].res.resOr.lpRes = lpsvcResLvl2;

			lpsvcResLvl2[0].rt = RES_CONTENT;
			lpsvcResLvl2[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
			lpsvcResLvl2[0].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
			lpsvcResLvl2[0].res.resContent.lpProp = lpSvcPropVal1;

			lpSvcPropVal1->ulPropTag = PR_SERVICE_NAME_A;
			lpSvcPropVal1->Value.lpszA = ConvertWideCharToMultiByte(L"MSPST MS");

			lpsvcResLvl2[1].rt = RES_CONTENT;
			lpsvcResLvl2[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING;
			lpsvcResLvl2[1].res.resContent.ulPropTag = PR_SERVICE_NAME_A;
			lpsvcResLvl2[1].res.resContent.lpProp = lpSvcPropVal2;

			lpSvcPropVal2->ulPropTag = PR_SERVICE_NAME_A;
			lpSvcPropVal2->Value.lpszA = ConvertWideCharToMultiByte(L"MSUPST MS");

			// Query the table to get the the default profile only
			CHK_HR_DBG(HrQueryAllRows(lpServiceTable,
				(LPSPropTagArray)& sptaSvcProps,
				lpSvcRes,
				NULL,
				0,
				&lpSvcRows), L"Calling HrQueryAllRows");

			if (lpSvcRows->cRows > 0)
			{
				wprintf(L"Found %i PST services in profile %s\n", lpSvcRows->cRows, lpszProfileName);
				// Start loop services
				for (unsigned int i = 0; i < lpSvcRows->cRows; i++)
				{

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
							(LPVOID*)& lpProvRes), L"Calling MAPIAllocateBuffer");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SRestriction) * 2,
							(LPVOID*)& lpProvResLvl1), L"Calling MAPIAllocateBuffer");

						CHK_HR_DBG(MAPIAllocateBuffer(
							sizeof(SPropValue),
							(LPVOID*)& lpProvPropVal), L"Calling MAPIAllocateBuffer");

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
							(LPSPropTagArray)& sptaProvProps,
							lpProvRes,
							NULL,
							0,
							&lpProvRows), L"Calling HrQueryAllRows");

						if (lpProvRows->cRows > 0)
						{

							LPPROFSECT lpProfSection = NULL;
							if (SUCCEEDED(lpServiceAdmin->OpenProfileSection((LPMAPIUID)lpProvRows->aRow->lpProps[iProvInstanceKey].Value.bin.lpb, NULL, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSection)))
							{
								LPMAPIPROP lpMAPIProp = NULL;
								if (SUCCEEDED(lpProfSection->QueryInterface(IID_IMAPIProp, (void**)& lpMAPIProp)))
								{
									LPSPropValue prDisplayName = NULL;
									if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_DISPLAY_NAME_W, &prDisplayName)))
									{
										// bind to the PR_PST_PATH_W property
										LPSPropValue pstPathW = NULL;
										if (SUCCEEDED(HrGetOneProp(lpMAPIProp, PR_PST_PATH_W, &pstPathW)))
										{
											if (pstPathW)
											{
												std::wstring szCurrentPath = ConvertWideCharToStdWstring(pstPathW->Value.lpszW);
												std::wstring szOldPath = szCurrentPath;
												LPWSTR lpszOldPath = (LPWSTR)szOldPath.c_str();
												if SUCCEEDED(PathRemoveFileSpec(lpszOldPath))
												{
													if (WStringReplace(&szCurrentPath, lpszOldPath, ConvertWideCharToStdWstring(lpszNewPath)))
													{
														if (bMoveFiles)
														{
															wprintf(L"Moving file %s to new location %s\n", pstPathW->Value.lpszW, szCurrentPath.c_str());
															BOOL bFileMoved = false;
															bFileMoved = MoveFileExW(pstPathW->Value.lpszW, (LPCWSTR)szCurrentPath.c_str(), MOVEFILE_COPY_ALLOWED | MOVEFILE_WRITE_THROUGH);
															if (bFileMoved)
															{
																wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
																pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
																CHK_HR_DBG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
															}
															else
															{
																wprintf(L"Unable to move file\n");
															}
														}
														else
														{
															wprintf(L"Updating path for data file named %s\n", pstPathW->Value.lpszW);
															pstPathW[0].Value.lpszW = (LPWSTR)szCurrentPath.c_str();
															CHK_HR_DBG(lpProfSection->SetProps(1, pstPathW, NULL), L"Calling SetProps");
														}
													}
												}
												if (pstPathW) MAPIFreeBuffer(pstPathW);
											}
										}
										if (prDisplayName) MAPIFreeBuffer(prDisplayName);
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
				if (lpSvcRows) FreeProws(lpSvcRows);
				// End loop services
			}
			if (lpSvcPropVal1) MAPIFreeBuffer(lpSvcPropVal1);
			if (lpSvcPropVal2) MAPIFreeBuffer(lpSvcPropVal2);
			if (lpsvcResLvl1) MAPIFreeBuffer(lpsvcResLvl1);
			if (lpsvcResLvl2) MAPIFreeBuffer(lpsvcResLvl2);
			if (lpSvcRes) MAPIFreeBuffer(lpSvcRes);
			if (lpServiceTable) lpServiceTable->Release();
			if (lpServiceAdmin) lpServiceAdmin->Release();
		}
		// End process services

		// Free up memory
		//if (lpProfPropVal) MAPIFreeBuffer(lpProfPropVal);
		//if (lpProfResLvl1) MAPIFreeBuffer(lpProfResLvl1);
		//if (lpProfRes) MAPIFreeBuffer(lpProfRes);
		if (lpProfRows) FreeProws(lpProfRows);
		if (lpProfTable) lpProfTable->Release();
		if (lpProfAdmin) lpProfAdmin->Release();

	Error:
		goto CleanUp;
	CleanUp:
		return hRes;
	}


	HRESULT HrCreatePstService(LPSERVICEADMIN2 lpServiceAdmin2, LPMAPIUID* lppServiceUid, LPWSTR lpszServiceName, ULONG ulResourceFlags, ULONG ulPstConfigFlag, LPWSTR lpszPstPathW, LPWSTR lpszDisplayName)
	{
		HRESULT			hRes = S_OK; // Result code returned from MAPI calls.
		SPropValue		rgvalStoreProvider[3];
		MAPIUID			uidService = { 0 };
		LPMAPIUID		lpServiceUid = &uidService;
		LPPROFSECT		lpProfSect = NULL;
		LPPROFSECT		lpStoreProviderSect = nullptr;
		LPMAPIPROP lpMapiProp = NULL;
		// Adds a message service to the current profile and returns that newly added service UID.
		CHK_HR_DBG(lpServiceAdmin2->CreateMsgServiceEx((LPTSTR)ConvertWideCharToMultiByte(lpszServiceName),
			(LPTSTR)ConvertWideCharToMultiByte(lpszDisplayName),
			NULL,
			0,
			&uidService), L"Calling CreateMsgServiceEx.");

		CHK_HR_DBG(lpServiceAdmin2->OpenProfileSection(&uidService,
			0,
			MAPI_FORCE_ACCESS | MAPI_MODIFY,
			&lpProfSect), L"Calling OpenProfileSection.");



		CHK_HR_DBG(lpProfSect->QueryInterface(IID_IMAPIProp, (LPVOID*)& lpMapiProp), L"Calling QueryInterface.");

		if (lpMapiProp)
		{
			LPSPropValue prResourceFlags;
			MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)& prResourceFlags);

			prResourceFlags->ulPropTag = PR_RESOURCE_FLAGS;
			prResourceFlags->Value.l = ulResourceFlags;
			CHK_HR_DBG(lpMapiProp->SetProps(1, prResourceFlags, NULL), L"Calling SetProps.");

			CHK_HR_DBG(lpMapiProp->SaveChanges(FORCE_SAVE), L"Calling SaveChanges.");
			MAPIFreeBuffer(prResourceFlags);
			lpMapiProp->Release();
		}

		MAPIAllocateBuffer(sizeof(LPPROFSECT), (LPVOID*)& lpStoreProviderSect);
		ZeroMemory(lpStoreProviderSect, sizeof(LPPROFSECT));

		CHK_HR_DBG(HrGetSections(lpServiceAdmin2, lpServiceUid, NULL, &lpStoreProviderSect), L"Calling HrGetSections.");

		// Set up a SPropValue array for the properties you need to configure.
		/*
		PR_PST_CONFIG_FLAGS
		PR_PST_PATH_W
		PR_DISPLAY_NAME_W
		*/

		ZeroMemory(&rgvalStoreProvider[0], sizeof(SPropValue));
		rgvalStoreProvider[0].ulPropTag = PR_PST_CONFIG_FLAGS;
		rgvalStoreProvider[0].Value.l = ulPstConfigFlag;

		ZeroMemory(&rgvalStoreProvider[1], sizeof(SPropValue));
		rgvalStoreProvider[1].ulPropTag = PR_PST_PATH_W;
		rgvalStoreProvider[1].Value.lpszW = lpszPstPathW;

		ZeroMemory(&rgvalStoreProvider[2], sizeof(SPropValue));
		rgvalStoreProvider[2].ulPropTag = PR_DISPLAY_NAME_W;
		rgvalStoreProvider[2].Value.lpszW = lpszDisplayName;

		CHK_HR_DBG(lpStoreProviderSect->SetProps(
			2,
			rgvalStoreProvider,
			nullptr), L"Calling SetProps.");

		CHK_HR_DBG(lpStoreProviderSect->SaveChanges(KEEP_OPEN_READWRITE), L"Calling SaveChanges.");
	Error:
		goto CleanUp;
	CleanUp:
		// Clean up
		if (lpStoreProviderSect) lpStoreProviderSect->Release();
		if (lpProfSect) lpProfSect->Release();
		return hRes;
	}
}