#pragma once
#include <string>
#include <Windows.h>
#include "Logger.h"
#include <sstream>
namespace MAPIToolkit
{
#define CHK_HR(_hRes) \
	do { \
		hRes = _hRes; \
		if (FAILED(_hRes)) \
																{ \
			std::cout << "FAILED! hr = " << std::hex << hRes << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			std::cout << " >>> " << std::hex << _hRes <<  "\n"; \
			goto Error; \
																} \
								} while (0)

#define HCK(_hRes) \
	do { \
		if (FAILED(_hRes)) \
																{ \
			std::cout << "FAILED! hr = " << std::hex << _hRes << ".  LINE = " << std::dec << __LINE__ << "\n"; \
			std::cout << " >>> " << std::hex << _hRes <<  "\n"; \
			goto Error; \
																} \
								} while (0)

#define HCKL(_hRes, loggerMode) \
	do { \
		if (FAILED(_hRes)) \
					{ \
			std::wostringstream oss; \
			oss << L"Error " << std::hex << _hRes << L" in file " << __FILE__ << L" at line " << std::dec << __LINE__ ; \
			Logger::WriteLine(LOGLEVEL_ERROR, oss.str()); \
			goto Error; \
																} \
								} while (0)

#define CHK_HR_MSG(_hRes, wszMessage) \
	do { \
		hRes = _hRes; \
		Logger::Write(LOGLEVEL_INFO, wszMessage); \
		if (FAILED(_hRes)) \
																		{ \
			std::wostringstream oss; \
			oss << L"Method: " << __FUNCTIONW__ << L"\nFile: " << __FILE__ << L"\nLine:  " << std::dec << __LINE__ << L"\nError: " << std::hex << _hRes ; \
			Logger::EndLine(LOGLEVEL_FAILED, L" ...FAIL"); \
			Logger::WriteLine(LOGLEVEL_ERROR, oss.str()); \
			goto Error; \
																		} \
			else \
				Logger::EndLine(LOGLEVEL_SUCCESS, L" ...SUCCESS"); \
									} while (0)

#define CHK_HR_DBG(_hRes, wszMessage) \
	do { \
			hRes = _hRes; \
			std::wostringstream oss; \
			if (FAILED(hRes)) \
			{ \
				oss << L"Error " << std::hex << hRes << L" at " << wszMessage << L" > " << __FUNCTIONW__; \
				Logger::WriteLine(LOGLEVEL_DEBUG, oss.str()); \
				goto Error; \
			} \
			else \
			{ \
				oss << wszMessage << L" > " << __FUNCTIONW__; \
				Logger::WriteLine(LOGLEVEL_DEBUG, oss.str()); \
			} \
		} while (0)

#define CHK_BOOL_MSG(boolVal, wszMessage) \
	do { \
		hRes = (true == boolVal) ? S_OK : S_FALSE; \
		Logger::Write(LOGLEVEL_INFO, wszMessage); \
		if (FAILED(hRes)) \
																		{ \
			std::wostringstream oss; \
			oss << L"Method: " << __FUNCTIONW__ << L"\nFile: " << __FILE__ << L"\nLine:  " << std::dec << __LINE__ << L"\nError: " << std::hex << hRes ; \
			Logger::EndLine(LOGLEVEL_FAILED, L" ...FAIL"); \
			Logger::WriteLine(LOGLEVEL_ERROR, oss.str()); \
			goto Error; \
																		} \
			else \
				Logger::EndLine(LOGLEVEL_SUCCESS, L" ...SUCCESS"); \
									} while (0)

#define FCHK(variable, flag) (flag == (variable & flag))

#define VCHK(variable, value) (value == variable)
}