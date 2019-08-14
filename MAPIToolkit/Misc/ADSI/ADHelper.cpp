#pragma once
#include "ADHelper.h"
#include "..//Utility/StringOperations.h"
#include "..//..//InlineAndMacros.h"

namespace MAPIToolkit
{
	// retrieves the current user DN
	std::wstring GetUserDn()
	{
		std::wstring wszUserDn = L"";
		HRESULT hRes = S_OK;

		IADsADSystemInfo* pADsys;
		HCK(CoCreateInstance(CLSID_ADSystemInfo,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_IADsADSystemInfo,
			(void**)& pADsys));

		if (pADsys)
		{
			BSTR bstrUserName = NULL;
			HCK(pADsys->get_UserName(&bstrUserName));
			if (bstrUserName)
			{
				wszUserDn = std::wstring(bstrUserName);
				::SysFreeString(bstrUserName);
			}
			pADsys->Release();
		}
	Error:
		return wszUserDn;
	}

	// Retrieves the DNS name of a DC in the local computer's domain
	std::wstring GetDCName()
	{
		std::wstring wszDCName = L"";
		HRESULT hRes = S_OK;

		IADsADSystemInfo* pADsys;
		HCK(CoCreateInstance(CLSID_ADSystemInfo,
			NULL,
			CLSCTX_INPROC_SERVER,
			IID_IADsADSystemInfo,
			(void**)& pADsys));

		if (pADsys)
		{
			BSTR bstrDCName = NULL;
			HCK(pADsys->GetAnyDCName(&bstrDCName));
			if (bstrDCName)
			{
				wszDCName = std::wstring(bstrDCName);
				::SysFreeString(bstrDCName);
			}
			pADsys->Release();
		}
	Error:
		return wszDCName;
	}

	// Find the primary smtp address of the current user
	// based on the information retrieved from the system
	std::wstring FindPrimarySMTPAddress(std::wstring wszUserDn)
	{
		HRESULT hr = S_OK;
		std::wstring wszSmtpAddress = L"";

		//Get rootDSE and the config container's DN.
		LPADS lpAds = NULL;
		wszUserDn = L"LDAP://" + wszUserDn;

		hr = ADsOpenObject((LPCWSTR)wszUserDn.c_str(),
			NULL,
			NULL,
			ADS_SECURE_AUTHENTICATION,
			//Use Secure Authentication
			IID_IADs,
			(void**)& lpAds);

		if ((S_OK == hr) && lpAds)
		{
			VARIANT varPropValue;
			BSTR bstrProperty = BSTR(L"proxyAddresses");
			hr = lpAds->Get(bstrProperty, &varPropValue);
			if ((SUCCEEDED(hr)) && (VT_VARIANT ^ varPropValue.vt))
			{
				LONG cElements, lLBound, lUBound;

				if (SafeArrayGetDim(varPropValue.parray) == 1)
				{
					// Get array bounds.
					hr = SafeArrayGetLBound(varPropValue.parray, 1, &lLBound);
					if (FAILED(hr))
						goto Error;
					hr = SafeArrayGetUBound(varPropValue.parray, 1, &lUBound);
					if (FAILED(hr))
						goto Error;

					cElements = lUBound - lLBound + 1;

					VARIANT propVal;
					VariantInit(&propVal);
					for (LONG i = 0; i < cElements - 1; i++)
					{
						hr = SafeArrayGetElement(varPropValue.parray, &i, &propVal);
						if (propVal.vt == VT_BSTR)
						{
							std::wstring wszAddress = std::wstring(propVal.bstrVal);
							size_t pos = wszAddress.find(L"SMTP:");
							if (pos != std::wstring::npos)
							{
								pos = wszAddress.find(L":");
								wszSmtpAddress = wszAddress.substr(pos + 1);
								break;
							}
						}
					}


				}
			}
			lpAds->Release();
		}
	Error:
		return wszSmtpAddress;
	}

	std::wstring LDAPSearchUserDN(std::wstring wszSearchAttributeName, std::wstring wszSearchValue, ULONG ulAdTimeout)
	{
		//Handle the command line arguments.
		std::wstring wszUserDN = L"";
		std::wstring wszAdsiPathName = L"";
		std::wstring wszSearchFilter = L"";
		HRESULT hRes = S_OK;

		//Get rootDSE and the config container's DN.
		IADs* pObject = NULL;
		IDirectorySearch* pConfigNC = NULL;
		VARIANT var;

		ADS_SEARCHPREF_INFO		prefs[5];
		ULONG					cPrefs = 0;

		// COL for iterations
		LPOLESTR pszColumn = NULL;
		ADS_SEARCH_COLUMN col;


		// Handle used for searching
		ADS_SEARCH_HANDLE hSearch = NULL;

		LPWSTR lpszAttrName = ConvertStdStringToWideChar(L"distinguishedName");
		int iCount = 0;
		DWORD x = 0L;

		HCK(ADsOpenObject(GetDCName().c_str(),
			NULL,
			NULL,
			ADS_SECURE_AUTHENTICATION,
			//Use Secure Authentication
			IID_IADs,
			(void**)& pObject));

		HCK(pObject->Get(ConvertStdStringToBstr(L"defaultNamingContext"), &var));

		wszAdsiPathName = L"LDAP://" + std::wstring(var.bstrVal);

		HCK(ADsOpenObject(wszAdsiPathName.c_str(),
			NULL,
			NULL,
			ADS_SECURE_AUTHENTICATION,
			//Use Secure Authentication
			IID_IDirectorySearch,
			(void**)& pConfigNC));

		wszSearchFilter = L"(&(objectcategory=user)(objectClass=person)(" + wszSearchAttributeName + L"=" + wszSearchValue + L"))";



		// Go find referrals
		prefs[cPrefs].dwSearchPref = ADS_SEARCHPREF_CHASE_REFERRALS;
		prefs[cPrefs].vValue.dwType = ADSTYPE_INTEGER;
		prefs[cPrefs].vValue.Integer = ADS_CHASE_REFERRALS_ALWAYS;
		cPrefs++;
		// Synchronous search please
		prefs[cPrefs].dwSearchPref = ADS_SEARCHPREF_ASYNCHRONOUS;
		prefs[cPrefs].vValue.dwType = ADSTYPE_BOOLEAN;
		prefs[cPrefs].vValue.Boolean = FALSE;
		cPrefs++;
		// Set timeouts to something reasonable
		prefs[cPrefs].dwSearchPref = ADS_SEARCHPREF_TIME_LIMIT;
		prefs[cPrefs].vValue.dwType = ADSTYPE_INTEGER;
		prefs[cPrefs].vValue.Integer = ulAdTimeout;
		cPrefs++;
		prefs[cPrefs].dwSearchPref = ADS_SEARCHPREF_TIMEOUT;
		prefs[cPrefs].vValue.dwType = ADSTYPE_INTEGER;
		prefs[cPrefs].vValue.Integer = ulAdTimeout + 1;
		cPrefs++;
		// Set scope to search entire subtree
		prefs[cPrefs].dwSearchPref = ADS_SEARCHPREF_SEARCH_SCOPE;
		prefs[cPrefs].vValue.dwType = ADSTYPE_INTEGER;
		prefs[cPrefs].vValue.Integer = ADS_SCOPE_SUBTREE;
		cPrefs++;



		// Set the search preference
		HCK(pConfigNC->SetSearchPreference(const_cast<ADS_SEARCHPREF_INFO*>(prefs), cPrefs));



		//Return non-verbose list properties only
		HCK(pConfigNC->ExecuteSearch((LPWSTR)wszSearchFilter.c_str(),
			&lpszAttrName,
			1,
			&hSearch));

		// Call IDirectorySearch::GetNextRow() to retrieve the next row 
		//of data
		HCK(pConfigNC->GetFirstRow(hSearch));

		while (hRes != S_ADS_NOMORE_ROWS)
		{
			//Keep track of count.
			iCount++;
			// loop through the array of passed column names,
			// print the data for each column
			while (pConfigNC->GetNextColumnName(hSearch, &pszColumn) != S_ADS_NOMORE_COLUMNS)
			{
				HCK(pConfigNC->GetColumn(hSearch, pszColumn, &col));
				if (0 == wcscmp(L"distinguishedName", pszColumn))
				{
					wszUserDN = std::wstring(col.pADsValues->CaseIgnoreString);
				}
				pConfigNC->FreeColumn(&col);
				FreeADsMem(pszColumn);
			}

			//Get the next row
			hRes = pConfigNC->GetNextRow(hSearch);
		}
		if (S_ADS_NOMORE_ROWS == hRes)
		{
			hRes = S_OK;
		}
		else
		{
			if (S_OK != hRes)
			{
				wprintf(L"An error occurred.\n  HRESULT: %x\n", hRes);
			}
		}

		// If facility is Win32, get the Win32 error 
		if (HRESULT_FACILITY(hRes) == FACILITY_WIN32)
		{
			DWORD dwLastError;
			WCHAR szErrorBuf[MAX_PATH];
			WCHAR szNameBuf[MAX_PATH];
			// Get extended error value.
			HRESULT hr_return = S_OK;
			hr_return = ADsGetLastError(&dwLastError,
				szErrorBuf,
				MAX_PATH,
				szNameBuf,
				MAX_PATH);
			if (SUCCEEDED(hr_return))
			{
				wprintf(L"Error Code: %d\n Error Text: %ws\n Provider: %ws\n", dwLastError, szErrorBuf, szNameBuf);
			}
		}


		// Close the search handle to clean up
		pConfigNC->CloseSearchHandle(hSearch);

		if (SUCCEEDED(hRes) && 0 == iCount)
			hRes = S_FALSE;
	Error:


		VariantClear(&var);
		if (pConfigNC) pConfigNC->Release();
		if (pObject) pObject->Release();

		return wszUserDN;
	}

}