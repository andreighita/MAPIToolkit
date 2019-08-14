#pragma once
#include <objbase.h>
#include <wchar.h>
#include <activeds.h>
#include <Iads.h>
#include <sddl.h>
#include <wchar.h>
#include <initguid.h>
#include <string>
#define USES_IID_IADsADSystemInfo
#define USES_IID_IDirectorySearch
#define USES_IID_IADs
typedef IADs FAR * LPADS;
typedef IDirectorySearch FAR * LPDIRECTORYSEARCH;

namespace MAPIToolkit
{
	std::wstring GetUserDn();
	std::wstring GetDCName();
	std::wstring LDAPSearchUserDN(std::wstring wszSearchAttributeName, std::wstring wszSearchValue, ULONG ulAdTimeout);
	std::wstring FindPrimarySMTPAddress(std::wstring wszUserDn);
}