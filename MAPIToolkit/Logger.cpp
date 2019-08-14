#pragma once
#include "Logger.h"
#include "ToolkitTypeDefs.h"
#include <sstream>
#pragma warning(disable:4996)

namespace MAPIToolkit
{
	std::wstring Logger::m_szLogFilePath;
	std::wofstream Logger::m_ofsLogFile;
	bool Logger::m_bIsLogFileOpen;
	ULONG Logger::m_loggingMode = LOGGINGMODE_CONSOLE;


	void Logger::Initialise(std::wstring wszPath)
	{
		Logger::m_szLogFilePath = wszPath;
		Logger::m_ofsLogFile = std::wofstream(Logger::m_szLogFilePath, std::ios_base::app);
		if (!Logger::m_ofsLogFile.is_open())
		{
			std::cerr << "Couldn't open 'output.txt'" << std::endl;
			Logger::m_bIsLogFileOpen = false;
		}
		else
			Logger::m_bIsLogFileOpen = true;
	}

	void Logger::SetLoggingMode(ULONG loggingMode)
	{
		Logger::m_loggingMode = loggingMode;
	}

	void Logger::SetFilePath(std::wstring wszFilePath)
	{
		Initialise(wszFilePath);
	}

	Logger::~Logger()
	{
		{
			if (m_bIsLogFileOpen)
			{
				m_ofsLogFile.close();
			}
		}
	}

	ULONG Logger::Write(ULONG llLevel, std::wstring szMessage, ULONG loggingMode)
	{
		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered" << L" ";
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			time_t  timev = time(0);
			struct tm* now = localtime(&timev);
			std::wstring wszTimeStampNow = std::to_wstring(now->tm_year + 1900) + L"." + std::to_wstring(now->tm_mon + 1) + L"." + std::to_wstring(now->tm_mday) + L" " + std::to_wstring(now->tm_hour) + L":" + std::to_wstring(now->tm_min) + L":" + std::to_wstring(now->tm_sec);

			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" INFO " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" INFO " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" WARNING " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" WARNING " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" ERROR " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" ERROR " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" SUCCESS " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" SUCCESS " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" FAILED " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" FAILED " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
						std::wcout << wszTimeStampNow << L" DEBUG " << szMessage << L" ";
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" UNKN " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" UNKN " << szMessage << L" ";
				}
				break;
			}

			if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << wszTimeStampNow << szMessage << L" ";
			else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL))
			{
				std::wcout << wszTimeStampNow << szMessage << L" ";
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << L" ";
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}

	ULONG Logger::Continue(ULONG llLevel, std::wstring szMessage, ULONG loggingMode)
	{
		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered" << L" ";
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile  << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout  << szMessage << L" ";
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
					std::wcout << szMessage << L" ";
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			}

			if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << szMessage << L" ";
			else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL))
			{
				std::wcout << szMessage << L" ";
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << L" ";
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}

	ULONG Logger::EndLine(ULONG llLevel, std::wstring szMessage, ULONG loggingMode)
	{
		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered" << std::endl;
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			}

			if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << szMessage << std::endl;
			else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL))
			{
				std::wcout << szMessage << std::endl;
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << std::endl;
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}


		ULONG Logger::WriteLine(ULONG llLevel, std::wstring szMessage, ULONG loggingMode)
	{
		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered";
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			time_t  timev = time(0);
			struct tm* now = localtime(&timev);
			std::wstring wszTimeStampNow = std::to_wstring(now->tm_year + 1900) + L"." + std::to_wstring(now->tm_mon + 1) + L"." + std::to_wstring(now->tm_mday) + L" " + std::to_wstring(now->tm_hour) + L":" + std::to_wstring(now->tm_min) + L":" + std::to_wstring(now->tm_sec);

			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" INFO " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" INFO " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" WARNING " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" WARNING " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" ERROR " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" ERROR " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" SUCCESS " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" SUCCESS " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" FAILED " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" FAILED " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
					std::wcout << wszTimeStampNow << L" DEBUG " << szMessage << std::endl;
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" UNKN " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" UNKN " << szMessage << std::endl;
				}
				break;
			}

			if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
				Logger::m_ofsLogFile << wszTimeStampNow << szMessage << std::endl;
			else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL))
			{
				std::wcout << wszTimeStampNow << szMessage << std::endl;
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << std::endl;
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}

	ULONG Logger::Write(ULONG llLevel, std::wstring szMessage)
	{
		ULONG loggingMode = Logger::m_loggingMode;

		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered" << L" ";
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			time_t  timev = time(0);

			struct tm* now = localtime(&timev);
			std::wstring wszTimeStampNow = std::to_wstring(now->tm_year + 1900) + L"." + std::to_wstring(now->tm_mon + 1) + L"." + std::to_wstring(now->tm_mday) + L" " + std::to_wstring(now->tm_hour) + L":" + std::to_wstring(now->tm_min) + L":" + std::to_wstring(now->tm_sec);

			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" INFO " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" INFO " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" WARNING " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" WARNING " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" ERROR " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" ERROR " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" SUCCESS " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" SUCCESS " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" FAILED " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" FAILED " << szMessage << L" ";
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
						std::wcout << wszTimeStampNow << L" DEBUG " << szMessage << L" ";;
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" UNKN " << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" UNKN " << szMessage << L" ";
				}
				break;
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << L" ";
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}

	ULONG Logger::Continue(ULONG llLevel, std::wstring szMessage)
	{
		ULONG loggingMode = Logger::m_loggingMode;

		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered" << std::endl;
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << L" ";
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
					std::wcout  << szMessage << L" ";;
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << L" ";
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout  << szMessage << L" ";
				}
				break;
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << L" ";
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}

	ULONG Logger::EndLine(ULONG llLevel, std::wstring szMessage)
	{
		ULONG loggingMode = Logger::m_loggingMode;

		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered" << L" ";
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << szMessage << std::endl;
				}
				break;
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << L" ";
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}

	ULONG Logger::WriteLine(ULONG llLevel, std::wstring szMessage)
	{
		ULONG loggingMode = Logger::m_loggingMode;

		if (loggingMode == LOGGINGMODE_NONE)
		{
			return LOGCALLSTATUS_LOGGINGDISABLED;
		}

		if (!Logger::m_szLogFilePath.empty() && !Logger::m_bIsLogFileOpen)
		{
			try
			{
				Logger::Initialise(Logger::m_szLogFilePath);
			}
			catch (int exception)
			{
				std::wostringstream oss; \
					oss << L"Error " << std::dec << exception << L" encountered" << std::endl;;
				Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
				return LOGCALLSTATUS_ERROR;
			}
		}
		try
		{
			time_t  timev = time(0);

			struct tm* now = localtime(&timev);
			std::wstring wszTimeStampNow = std::to_wstring(now->tm_year + 1900) + L"." + std::to_wstring(now->tm_mon + 1) + L"." + std::to_wstring(now->tm_mday) + L" " + std::to_wstring(now->tm_hour) + L":" + std::to_wstring(now->tm_min) + L":" + std::to_wstring(now->tm_sec);

			switch (llLevel)
			{
			case LOGLEVEL_INFO:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" INFO " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" INFO " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_WARNING:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" WARNING " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" WARNING " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_ERROR:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" ERROR " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" ERROR " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_SUCCESS:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" SUCCESS " << szMessage << std::endl;
else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" SUCCESS " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_FAILED:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" FAILED " << szMessage << std::endl;
				else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" FAILED " << szMessage << std::endl;
				}
				break;
			case LOGLEVEL_DEBUG:
				if (loggingMode == LOGGINGMODE_DEBUG)
				{
						std::wcout << wszTimeStampNow << L" DEBUG " << szMessage << std::endl;;
				}
				break;
			default:
				if ((loggingMode == LOGGINGMODE_DEBUG || loggingMode == LOGGINGMODE_ALL || loggingMode == LOGGINGMODE_FILE) && Logger::m_bIsLogFileOpen)
					Logger::m_ofsLogFile << wszTimeStampNow << L" UNKN " << szMessage << std::endl;
				else if ((loggingMode == LOGGINGMODE_CONSOLE) || (loggingMode == LOGGINGMODE_ALL) || (loggingMode == LOGGINGMODE_DEBUG))
				{
					std::wcout << wszTimeStampNow << L" UNKN " << szMessage << std::endl;
				}
				break;
			}
			return LOGCALLSTATUS_SUCCESS;
		}
		catch (int exception)
		{
			std::wostringstream oss; \
				oss << L"Error " << std::dec << exception << L" encountered" << std::endl;;
			Logger::Write(LOGLEVEL_ERROR, oss.str(), loggingMode);
			return LOGCALLSTATUS_ERROR;
		}
		return LOGCALLSTATUS_ERROR;

	}
}