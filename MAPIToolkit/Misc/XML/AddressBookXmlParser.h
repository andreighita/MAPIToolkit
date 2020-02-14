#pragma once
#include <windows.h>
#include <msxml.h>
#include <objsafe.h>
#include <objbase.h>
#include <atlbase.h>
#include "wchar.h"
#include "../../ToolkitTypeDefs.h"
#include "../../InlineAndMacros.h"
#include "../../Toolkit.h"

#include <stdio.h>
#include <tchar.h>
#include <windows.h>
#include <comdef.h>
#import <msxml6.dll> rename_namespace(_T("MSXML"))

namespace MAPIToolkit
{
	HRESULT ParseAddressBookXml(LPTSTR lpszABConfigurationPath)
	{
		if (!lpszABConfigurationPath)
		{
			return E_FAIL;
		}

		HRESULT hr = S_OK;
		VARIANT_BOOL bSuccess = false;
		MSXML::IXMLDOMDocument2Ptr pXMLDoc;
		try {
			hr = pXMLDoc.CreateInstance(__uuidof(MSXML::DOMDocument60),
				NULL, CLSCTX_INPROC_SERVER);

		}
		catch (_exception& e)
		{

		}

		bSuccess = pXMLDoc->load(lpszABConfigurationPath);
		DWORD errr = GetLastError();
		if (bSuccess)
		{
			// Get a pointer to the root
			MSXML::IXMLDOMElement* pRootElm;
			HCK(pXMLDoc->get_documentElement(&pRootElm));

			MSXML::IXMLDOMNodeList* pXMLNodes;
			// Thanks to the magic of CComPtr, we never need call
			// Release() -- that gets done automatically.
			HCK(pRootElm->get_childNodes(&pXMLNodes));
			long lCount;
			HCK(pXMLNodes->get_length(&lCount));
			for (int i = 0; i < lCount; i++)
			{
				MSXML::IXMLDOMNode* pXMLNode;
				HCK(pXMLNodes->get_item(i, &pXMLNode));
				BSTR bstrNodeName;
				HCK(pXMLNode->get_nodeName(&bstrNodeName));
				BSTR bstrNodeText;
				HCK(pXMLNode->get_text(&bstrNodeText));

				try
				{
					if (Toolkit::m_preserveCase == TRUE)
						Toolkit::g_addressBookMap.at(StringToLower(bstrNodeName)) = bstrNodeText;
					else
						Toolkit::g_addressBookMap.at(StringToLower(bstrNodeName)) = StringToLower(bstrNodeText);
				}
				catch (const std::exception& e)
				{

				}

			}
		}

	Error:
		return hr;
	}
}