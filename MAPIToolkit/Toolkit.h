#pragma once
#include <Windows.h>
#include <tchar.h>
#include "ToolkitTypeDefs.h"


#include <MAPIX.h>
namespace MAPIToolkit
{
	class Toolkit
	{

	protected:
		static void DisplayUsage();
		static BOOL Is64BitProcess(void);
		static int GetOutlookVersion();
		static BOOL IsCorrectBitness();

		static VOID RunAction();
		static BOOL ParseParams(int argc, wchar_t* argv[]);
		static BOOL SaveConfig();
		static BOOL ReadConfig();
		static BOOL Initialise();
		static VOID Uninitialise();

		

		static std::map<std::wstring, ULONG> g_actionsMap;
		static std::map<std::wstring, ULONG> g_profileModeMap;
		static std::map<std::wstring, ULONG> g_serviceModeMap;
		static std::map<std::wstring, ULONG> g_serviceTypeMap;
		static std::map<std::wstring, ULONG> g_loggingModeMap;

		static ULONG m_action;
		//static int m_OutlookVersion;
		//static ULONG m_loggingMode;
		//static ServiceWorker* m_serviceWorker;
		//static ProviderWorker* m_providerWorker;
		//static ProfileWorker* m_profileWorker;
		//static ULONG m_profileCount;
		//static std::wstring m_wszExportPath;
		//static ULONG m_exportMode; // 0 = no export; 1 = export;
		//static std::wstring m_wszLogFilePath;
		//static ULONG m_profileMode; // pm
		static LPPROFADMIN m_pProfAdmin;
		//static ULONG m_serviceType; // pm
		static BOOL m_registry;
	public:
		static std::map<std::wstring, std::wstring> g_addressBookMap;
		static std::map<std::wstring, std::wstring> g_toolkitMap;
		static std::map<std::wstring, std::wstring> g_regKeyMap;
		static std::map<int, std::wstring > g_hexMap;
		static std::map<std::wstring, ULONG> g_saveConfigMap;
		static std::map<std::wstring, std::wstring> g_parameterHelpInfo;
		static std::map <std::wstring, std::wstring > g_parameterHelpValues;
	public:
		 static VOID Run(int argc, wchar_t* argv[]);

	private:
		static BOOL RunActionOneProfile(std::wstring wszProfileName);
		static BOOL RunActionOneService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID pMapiUid);

		// ACTION_SERVICE_ADD	
		static VOID AddService(LPSERVICEADMIN2 pServiceAdmin);
		static BOOL AddAddressBookService(LPSERVICEADMIN2 pServiceAdmin);

		// ACTION_SERVICE_UPDATE	
		static VOID UpdateService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID lpMAPIUid);
		static VOID UpdateAddressBookService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID lpMAPIUid);

		// ACTION_SERVICE_LIST
		static VOID ListService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID lpMAPIUid);

		// ACTION_SERVICE_LISTALL
		static VOID ListAllServices(LPSERVICEADMIN2 pServiceAdmin);

		// ACTION_SERVICE_REMOVE
		static VOID RemoveService(LPSERVICEADMIN2 pServiceAdmin, LPMAPIUID pMapiUid);

		// ACTION_SERVICE_REMOVEALL
		static VOID RemoveAllServices(LPSERVICEADMIN2 pServiceAdmin);
	};
}