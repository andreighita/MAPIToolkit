#pragma once

#include <iostream>
#include <fstream>
#include <string>
#include <time.h>
#include "ToolkitTypeDefs.h"
// Log, version 0.1: a simple logging class
namespace MAPIToolkit
{
	class Logger
	{
	public:
		static void Initialise(std::wstring wszPath);
		static ULONG Write(ULONG llLevel, std::wstring szMessage, ULONG loggingMode);
		static ULONG Write(ULONG llLevel, std::wstring szMessage);
		static ULONG Continue(ULONG llLevel, std::wstring szMessage, ULONG loggingMode);
		static ULONG Continue(ULONG llLevel, std::wstring szMessage);
		static ULONG EndLine(ULONG llLevel, std::wstring szMessage, ULONG loggingMode);
		static ULONG EndLine(ULONG llLevel, std::wstring szMessage);
		static ULONG WriteLine(ULONG llLevel, std::wstring szMessage, ULONG loggingMode);
		static ULONG WriteLine(ULONG llLevel, std::wstring szMessage);
		static void SetLoggingMode(ULONG loggingMode);
		static void SetFilePath(std::wstring wszFilePath);
	private:
		~Logger();

		static std::wofstream m_ofsLogFile;
		static std::wstring m_szLogFilePath;
		static bool m_bIsLogFileOpen;
		static ULONG m_loggingMode;
		ULONG m_logLevel;

	};
}
