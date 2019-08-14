#pragma once
#include <stdio.h>
#include <tchar.h>
#include <msxml6.h>
#include <Windows.h>
#include <string.h>
#include "..//../ToolkitTypeDefs.h"

// Macro to verify memory allcation.
#define CHK_ALLOC(p)        do { if (!(p)) { hr = E_OUTOFMEMORY; goto CleanUp; } } while(0)

// Macro that releases a COM object if not NULL.
#define SAFE_RELEASE(p)     do { if ((p)) { (p)->Release(); (p) = NULL; } } while(0)

namespace MAPIToolkit
{
	// Helper function to create a VT_BSTR variant from a null terminated string. 
	HRESULT VariantFromString(PCWSTR wszValue, VARIANT& Variant);

	// Helper function to create a DOM instance. 
	HRESULT CreateAndInitDOM(IXMLDOMDocument** ppDoc);

	// Helper that allocates the BSTR param for the caller.
	HRESULT CreateElement(IXMLDOMDocument* pXMLDom, PCWSTR wszName, IXMLDOMElement** ppElement);

	// Helper function to append a child to a parent node.
	HRESULT AppendChildToParent(IXMLDOMNode* pChild, IXMLDOMNode* pParent);

	// Helper function to create and add a processing instruction to a document node.
	HRESULT CreateAndAddPINode(IXMLDOMDocument* pDom, PCWSTR wszTarget, PCWSTR wszData);

	// Helper function to create and add a comment to a document node.
	HRESULT CreateAndAddCommentNode(IXMLDOMDocument* pDom, PCWSTR wszComment);

	// Helper function to create and add an attribute to a parent node.
	HRESULT CreateAndAddAttributeNode(IXMLDOMDocument* pDom, PCWSTR wszName, PCWSTR wszValue, IXMLDOMElement* pParent);

	// Helper function to create and append a text node to a parent node.
	HRESULT CreateAndAddTextNode(IXMLDOMDocument* pDom, PCWSTR wszText, IXMLDOMNode* pParent);
	// Helper function to create and append a CDATA node to a parent node.
	HRESULT CreateAndAddCDATANode(IXMLDOMDocument* pDom, PCWSTR wszCDATA, IXMLDOMNode* pParent);

	// Helper function to create and append an element node to a parent node, and pass the newly created
	// element node to caller if it wants.
	HRESULT CreateAndAddElementNode(IXMLDOMDocument* pDom, PCWSTR wszName, PCWSTR wszNewline, IXMLDOMNode* pParent, IXMLDOMElement** ppElement);

	void ExportXML(ULONG cProfileInfo, ProfileInfo* profileInfo, std::wstring szExportPath);

}