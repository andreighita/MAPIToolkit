#pragma once
// QueryKey - Enumerates the subkeys of key and its associated values.
//     hKey - Key whose subkeys and values are to be enumerated.
#include "RegistryHelper.h"
#include <windows.h>

#include <string>
#include <tchar.h>
#include "Misc/Utility/StringOperations.h"
#include "Toolkit.h"
#include "InlineAndMacros.h"

#define MAX_KEY_LENGTH 255
#define MAX_VALUE_NAME 16383
namespace MAPIToolkit
{
	void QueryKey(HKEY hKey)
	{
		TCHAR    achKey[MAX_KEY_LENGTH];   // buffer for subkey name
		DWORD    cbName;                   // size of name string 
		TCHAR    achClass[MAX_PATH] = TEXT("");  // buffer for class name 
		DWORD    cchClassName = MAX_PATH;  // size of class string 
		DWORD    cSubKeys = 0;               // number of subkeys 
		DWORD    cbMaxSubKey;              // longest subkey size 
		DWORD    cchMaxClass;              // longest class string 
		DWORD    cValues;              // number of values for key 
		DWORD    cchMaxValue;          // longest value name 
		DWORD    cbMaxValueData;       // longest value data 
		DWORD    cbSecurityDescriptor; // size of security descriptor 
		FILETIME ftLastWriteTime;      // last write time 

		DWORD i, retCode;

		TCHAR  achValue[MAX_VALUE_NAME];
		DWORD cchValue = MAX_VALUE_NAME;

		// Get the class name and the value count. 
		retCode = RegQueryInfoKey(
			hKey,                    // key handle 
			achClass,                // buffer for class name 
			&cchClassName,           // size of class string 
			NULL,                    // reserved 
			&cSubKeys,               // number of subkeys 
			&cbMaxSubKey,            // longest subkey size 
			&cchMaxClass,            // longest class string 
			&cValues,                // number of values for this key 
			&cchMaxValue,            // longest value name 
			&cbMaxValueData,         // longest value data 
			&cbSecurityDescriptor,   // security descriptor 
			&ftLastWriteTime);       // last write time 

									 // Enumerate the subkeys, until RegEnumKeyEx fails.

		if (cSubKeys)
		{
			wprintf(L"\nNumber of subkeys: %d\n", cSubKeys);

			for (i = 0; i < cSubKeys; i++)
			{
				cbName = MAX_KEY_LENGTH;
				retCode = RegEnumKeyEx(hKey, i,
					achKey,
					&cbName,
					NULL,
					NULL,
					NULL,
					&ftLastWriteTime);
				if (retCode == ERROR_SUCCESS)
				{
					_tprintf(TEXT("(%d) %s\n"), i + 1, achKey);
				}
			}
		}

		// Enumerate the key values. 

		if (cValues)
		{
			wprintf(L"\nNumber of values: %d\n", cValues);

			for (i = 0, retCode = ERROR_SUCCESS; i < cValues; i++)
			{
				cchValue = MAX_VALUE_NAME;
				achValue[0] = '\0';
				retCode = RegEnumValue(hKey, i,
					achValue,
					&cchValue,
					NULL,
					NULL,
					NULL,
					NULL);

				if (retCode == ERROR_SUCCESS)
				{
					_tprintf(TEXT("(%d) %s\n"), i + 1, achValue);
				}
			}
		}
	}

	bool __cdecl GetValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, REGSAM regSam, DWORD dwSearchedType, LPBYTE* lpszValueData, DWORD* dwValueDataSize)
	{
		HKEY hKey;
		bool fFound = false;
		LSTATUS lStatus = ERROR_SUCCESS;
		lStatus = RegOpenKeyEx(hRegistryHive,
			lpszKeyName,
			0,
			regSam,
			&hKey);
		if (lStatus == ERROR_SUCCESS
			)
		{
			DWORD dwType;
			DWORD dwSize;
			LPBYTE lpbLookupDataValue = NULL;

			if (ERROR_SUCCESS == RegQueryValueExW(hKey, lpszValueName, NULL, &dwType, NULL, &dwSize))
			{
				if (dwSearchedType == dwType)
				{
					lpbLookupDataValue = (LPBYTE)malloc(dwSize);
					ZeroMemory(lpbLookupDataValue, dwSize);
					if (ERROR_SUCCESS == RegQueryValueExW(hKey, lpszValueName, NULL, &dwType, lpbLookupDataValue, &dwSize))
					{
						memcpy(dwValueDataSize, &dwSize, sizeof(DWORD));
						memcpy(lpszValueData, &lpbLookupDataValue, sizeof(&lpbLookupDataValue));
						fFound = true;
					}
				}
			}
		}

		RegCloseKey(hKey);
		return fFound;
	}

	std::wstring __cdecl GetRegStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName)
	{
		LPBYTE lpbTempValueData = NULL;
		DWORD dwValueDataSize;

		if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ, REG_SZ, &lpbTempValueData, &dwValueDataSize))
		{
			if (lpbTempValueData)
			{
				return std::wstring((LPWSTR)lpbTempValueData);
			}
			else return L"";
		}
		else if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ | KEY_WOW64_64KEY, REG_SZ, &lpbTempValueData, &dwValueDataSize))
		{
			if (lpbTempValueData)
			{
				return std::wstring((LPWSTR)lpbTempValueData);
			}
			else return L"";
		}
		if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ | KEY_WOW64_32KEY, REG_SZ, &lpbTempValueData, &dwValueDataSize))
		{
			if (lpbTempValueData)
			{
				return std::wstring((LPWSTR)lpbTempValueData);
			}
			else return L"";
		}
		else
			return L"";
	}

	DWORD __cdecl GetRegDwordValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName)
	{
		DWORD* pdwTempValueData = NULL;
		DWORD dwValueDataSize;

		if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ, REG_DWORD, (LPBYTE*)& pdwTempValueData, &dwValueDataSize))
		{
			if (pdwTempValueData)
			{
				return *pdwTempValueData;
			}
			else return 0;
		}
		else if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ | KEY_WOW64_64KEY, REG_SZ, (LPBYTE*)& pdwTempValueData, &dwValueDataSize))
		{
			if (pdwTempValueData)
			{
				return *pdwTempValueData;
			}
			else return 0;
		}
		if (GetValue(hRegistryHive, lpszKeyName, lpszValueName, KEY_READ | KEY_WOW64_32KEY, REG_SZ, (LPBYTE*)& pdwTempValueData, &dwValueDataSize))
		{
			if (pdwTempValueData)
			{
				return *pdwTempValueData;
			}
			else return 0;
		}
		else
			return 0;
	}

	BOOL __cdecl WriteRegStringValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, LPCTSTR lpszValueData)
	{
		HKEY key;
		HRESULT hres = RegOpenKey(hRegistryHive, lpszKeyName, &key);
		if (hres != ERROR_SUCCESS)
		{
			if (hres == ERROR_FILE_NOT_FOUND)
			{
				hres = RegCreateKey(hRegistryHive, lpszKeyName, &key);
				if (hres != ERROR_SUCCESS)
				{
					return FALSE;
				}
			}
			else
			{
				Logger::Write(LOGLEVEL_ERROR, L"Unable to open registry key");
				return FALSE;
			}
		}

		if (RegSetValueExW(key, lpszValueName, 0, REG_SZ, (LPBYTE)lpszValueData, wcslen(lpszValueData) * sizeof(wchar_t)) != ERROR_SUCCESS)
		{
			RegCloseKey(key);
			Logger::Write(LOGLEVEL_ERROR, L"Unable to set registry value");
			return FALSE;
		}

		//Logger::Write(LOGLEVEL_SUCCESS, L"value  was set");
		RegCloseKey(key);
		return TRUE;
	}

	BOOL __cdecl WriteRegDwordValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, DWORD dwValueData)
	{
		HKEY key;
		HRESULT hres = RegOpenKey(hRegistryHive, lpszKeyName, &key);
		if (hres != ERROR_SUCCESS)
		{
			if (hres == ERROR_FILE_NOT_FOUND)
			{
				hres = RegCreateKey(hRegistryHive, lpszKeyName, &key);
				if (hres != ERROR_SUCCESS)
				{
					return FALSE;
				}
			}
			else
			{
				Logger::Write(LOGLEVEL_ERROR, L"Unable to open registry key");
				return FALSE;
			}
		}

		if (RegSetValueExW(key, lpszValueName, 0, REG_DWORD, (LPBYTE)& dwValueData, sizeof(dwValueData)) != ERROR_SUCCESS)
		{
			RegCloseKey(key);
			Logger::Write(LOGLEVEL_ERROR, L"Unable to set registry value value_name");
			return FALSE;
		}

		Logger::Write(LOGLEVEL_SUCCESS, L"value  was set");
		RegCloseKey(key);
		return TRUE;
	}

	BOOL __cdecl WriteRegBinaryValue(HKEY hRegistryHive, LPCTSTR lpszKeyName, LPCTSTR lpszValueName, BYTE* pbValueData)
	{
		HKEY key;
		HRESULT hres = RegOpenKey(hRegistryHive, lpszKeyName, &key);
		if (hres != ERROR_SUCCESS)
		{
			if (hres == ERROR_FILE_NOT_FOUND)
			{
				hres = RegCreateKey(hRegistryHive, lpszKeyName, &key);
				if (hres != ERROR_SUCCESS)
				{
					return FALSE;
				}
			}
			else
			{
				Logger::Write(LOGLEVEL_ERROR, L"Unable to open registry key");
				return FALSE;
			}
		}

		if (RegSetValueExW(key, lpszValueName, 0, REG_BINARY, pbValueData, sizeof(pbValueData)) != ERROR_SUCCESS)
		{
			RegCloseKey(key);
			Logger::Write(LOGLEVEL_ERROR, L"Unable to set registry value value_name");
			return FALSE;
		}

		Logger::Write(LOGLEVEL_SUCCESS, L"value  was set");
		RegCloseKey(key);
		return TRUE;
	}

	BOOL ReadRegConfig(HKEY hRegistryHive, LPCTSTR lpszKeyName, std::map<std::wstring, std::wstring>* pMapObject)
	{
		HRESULT hRes = S_OK;
		HKEY hKey;
		DWORD dwType = REG_SZ;
		LPDWORD lpdwType = &dwType;
		CHK_HR_DBG(RegOpenKey(hRegistryHive, lpszKeyName, &hKey), L"RegOpenKey");

		DWORD    cValues;              // number of values for key 
		DWORD    cbMaxValueData;       // longest value data 

		DWORD i, retCode;

		// Get the class name and the value count. 
		CHK_HR_DBG(RegQueryInfoKey(
			hKey,                    // key handle 
			NULL,                // buffer for class name 
			NULL,           // size of class string 
			NULL,                    // reserved 
			NULL,               // number of subkeys 
			NULL,            // longest subkey size 
			NULL,            // longest class string 
			&cValues,                // number of values for this key 
			NULL,            // longest value name 
			&cbMaxValueData,         // longest value data 
			NULL,   // security descriptor 
			NULL),     // last write time 
			L"Getting registry information");


		if (cValues)
		{
			for (i = 0, retCode = ERROR_SUCCESS; i < cValues; i++)
			{
				DWORD dwcValName = MAX_VALUE_NAME;
				LPDWORD lpdcValName = &dwcValName;
				LPWSTR lpRegValue = NULL;
				lpRegValue = (LPWSTR)malloc(MAX_VALUE_NAME);

				CHK_HR_DBG(RegEnumValueW(hKey, i,
					lpRegValue,
					lpdcValName,
					NULL,
					lpdwType,
					NULL,
					NULL), L"Getting reg value names");

					DWORD dwcValData = MAX_PATH;
					LPDWORD lpdwcValData = &dwcValData;
					LPWSTR lpRegValueData = NULL;
					lpRegValueData = (LPWSTR)malloc(MAX_PATH);

					CHK_HR_DBG(RegQueryValueEx(hKey, lpRegValue, 0, NULL, (LPBYTE)lpRegValueData, lpdwcValData), L"RegQueryValueEx");
					std::wstring wszValue = ConvertWideCharToStdWstring(lpRegValue);
					std::wstring wszValueData = ConvertWideCharToStdWstring(lpRegValueData);
					try
					{
						pMapObject->at(wszValue) = wszValueData;
					}
					catch (const std::exception& e)
					{

					}
					MAPIFreeBuffer(lpdwcValData);
					MAPIFreeBuffer(lpRegValueData);
				
				MAPIFreeBuffer(lpdcValName);
				MAPIFreeBuffer(lpRegValue);
			}
		}
		RegCloseKey(hKey);
	Error:
		goto CleanUp;
	CleanUp:
		return true;
	}

}
