/*
  Copyright (C) 2005-2008 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

#include "stdafx.h"
#include "DetectFx.h"
#include "ExcelClrLoader.h"

typedef HRESULT (STDAPICALLTYPE *pfnCorBindToRuntimeEx)(        
									LPWSTR pwszVersion,   
									LPWSTR pwszBuildFlavor, 
									DWORD flags,            
									REFCLSID rclsid,      
									REFIID riid,    
									LPVOID* ppv );


typedef HRESULT (STDAPICALLTYPE *pfnCorBindToRuntimeByCfg)(        
									IStream    *pCfgStream, 
									DWORD      reserved, 
									DWORD      startupFlags, 
									REFCLSID   rclsid,
									REFIID     riid, 
									LPVOID FAR* ppv );

typedef HRESULT (STDAPICALLTYPE *pfnCorBindToCurrentRuntime)(        
									LPCWSTR pwszFileName,   
									REFCLSID rclsid,      
									REFIID riid,    
									LPVOID *ppv );


typedef HRESULT (STDAPICALLTYPE *pfnGetCORVersion)(
										LPWSTR szBuffer, 
                                        DWORD cchBuffer,
                                        DWORD* dwLength);


typedef HRESULT (STDAPICALLTYPE *pfnGetVersionFromProcess)(
										 HANDLE hProcess,
										 LPWSTR pBuffer, 
                                         DWORD cchBuffer,
                                         DWORD* dwLength);

HRESULT ExcelClrLoadv2(ICorRuntimeHost **ppHost);
//HRESULT ExcelClrLoadByConfig(ICorRuntimeHost **ppHost);
//HRESULT ExcelClrLoadByTempConfigFile(ICorRuntimeHost** ppHost);
//HRESULT ExcelClrLoadByConfigFile(LPCWSTR szConfigFileName, ICorRuntimeHost** ppHost);
HRESULT RepairRegistryAppPatch();

//// Loads version 2 or later of the CLR into Excel, and starts the runtime
//// returns S_FALSE if the Clr was already loaded
//// returns S_OK if the Clr is now loaded and started
//// returns E_FAIL if it can't, e.g. if the wrong version is already loaded into the process.
//// Priority is to get it done, however sneakily, if possible.
//// Does registry fiddling if possible, does config faking if needed, can still fail....
//// Shows messages detailing errors and how to fix.
//HRESULT ExcelClrLoad(ICorRuntimeHost **ppHost)
//{
//	HRESULT hr = E_FAIL;
//
//	// First attempt - just load the newest version of the Clr
//	HMODULE hMscoree = NULL;
//	hMscoree = LoadLibraryA("mscoree.dll");
//	if (hMscoree == 0)
//	{
//		// Should not happen - we have already checked that .Net 2.0 is installed
//		MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//						 "There was a problem while loading the .Net runtime:\r\n"
//						 "\tExcelClrLoader - Could not load library mscoree.dll.\r\n\r\n"
//						 "This is an unexpected catastrophic failure.", 
//						 "ExcelDna Add-In Loader", 0);
//		hr = E_FAIL;
//	}
//	else
//	{
//		
//		FreeLibrary(hMscoree);
//	}
//	return hr;
//
//	return S_OK;
//}

//// Used as a fallback if the registry cannot be edited.
//// TODO: Fill in newest version of the runtime, rather than fixed version 2.0.
//HRESULT ExcelClrLoadByConfig(ICorRuntimeHost **ppHost)
//{
//	HRESULT hr = E_FAIL;
//
//	HGLOBAL hConfigAlloc = GlobalAlloc(GHND, 100);
//	CHAR szConfig[100] =	"<configuration>"
//							"<startup>"
//							"<requiredRuntime version=\"v2.0.50727\" />"
//							"</startup></configuration>";
//	LPVOID pConfig = GlobalLock(hConfigAlloc);
//	memcpy(pConfig, szConfig, 100);
//	GlobalUnlock(hConfigAlloc);
//
//	IStream *pConfigStream = NULL;
//	hr = CreateStreamOnHGlobal(hConfigAlloc, FALSE, &pConfigStream);
//	if (FAILED(hr))
//	{
//		MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//						 "There was a problem while loading the .Net runtime:\r\n"
//						 "\tExcelClrLoadByConfig - Could not create stream for configuration.\r\n\r\n"
//						 "This is an unexpected catastrophic failure.", 
//						 "ExcelDna Add-In Loader", 0);
//		GlobalFree(hConfigAlloc);
//		DebugBreak();
//		hr = E_FAIL;
//	}
//	else
//	{
//		ULARGE_INTEGER size;
//		size.QuadPart = 90ui64;
//		LARGE_INTEGER zero;
//		zero.QuadPart = 0i64;
//
//		hr = pConfigStream->SetSize(size);
//		hr = pConfigStream->Seek(zero, STREAM_SEEK_SET, NULL);
//		// First attempt - just load the newest version of the Clr
//		HMODULE hMscoree = NULL;
//		hMscoree = LoadLibraryA("mscoree.dll");
//		if (hMscoree == 0)
//		{
//			MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//							 "There was a problem while loading the .Net runtime:\r\n"
//							 "\tExcelClrLoader - Could not load library mscoree.dll.\r\n\r\n"
//							 "This is an unexpected catastrophic failure.", 
//							 "ExcelDna Add-In Loader", 0);
//			hr = E_FAIL;					
//		}
//		else
//		{
//			pfnCorBindToRuntimeByCfg CorBindToRuntimeByCfg = (pfnCorBindToRuntimeByCfg)GetProcAddress(hMscoree, "CorBindToRuntimeByCfg");
//			if (CorBindToRuntimeByCfg == 0)
//			{
//				MessageBox(NULL,"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//								"There was a problem while loading the .Net runtime:\r\n"
//								"\tExcelClrLoader - Could not get proc address for CorBindToRuntimeByCfg.\r\n\r\n"
//								"This is an unexpected catastrophic failure.", 
//								"ExcelDna Add-In Loader", 0);
//				hr = E_FAIL;
//			}
//			else
//			{
//				hr = CorBindToRuntimeByCfg( pConfigStream, 0, 0, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (LPVOID*)ppHost);
//				if (FAILED(hr))
//				{
//					MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//									 "There was a problem while loading the .Net runtime:\r\n"
//									 "\tExcelClrLoader - CorBindToRuntimeByCfg failed.\r\n"
//									 "This error can occur if the .Net runtime is not installed, if version 2 or later is not installed, if the AppPatch registry entries are incorrect, or if there is a config file for Excel retricting the versions of the .Net runtime that can be loaded.\r\n\r\n"
//									 "Ensure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", 
//									 "ExcelDna Add-In Loader", 0);
//					hr = E_FAIL;
//				}
//				else
//				{
//					// Check the version that was loaded.
//					pfnGetCORVersion GetCORVersion = (pfnGetCORVersion)GetProcAddress(hMscoree, "GetCORVersion");
//					if (GetCORVersion == 0)
//					{
//						MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//										 "There was a problem while loading the .Net runtime:\r\n"
//										 "\tExcelClrLoadByConfig - Could not get proc address for GetCORVersion.\r\n"
//										 "This is an unknown, catastrophic failure.", 
//										 "ExcelDna Add-In Loader", 0);
//						hr = E_FAIL;
//					}
//					else
//					{
//						// Display current runtime loaded, or that will be loaded:
//						WCHAR szVersion[MAX_PATH + 1];
//						DWORD dwLength = MAX_PATH;
//
//						hr = GetCORVersion(szVersion, dwLength, &dwLength);
//						if (FAILED(hr))
//						{
//							MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//											 "There was a problem while loading the .Net runtime:\r\n"
//											 "\tExcelClrLoadByConfig - GetCORVersion failed.\r\n"
//											 "This is an unknown, catastrophic failure.", 
//											 "ExcelDna Add-In Loader", 0);
//							hr = E_FAIL;
//						}
//						else
//						{
//#ifdef _DEBUG
//							MessageBoxW(NULL, L"Version after loading:" , L"ExcelDna Add-In Loader", 0);
//							MessageBoxW(NULL, szVersion , L"ExcelDna Add-In Loader", 0);
//#endif
//							// TODO: Proper parse of version etc.
//							if ( wcsncmp(szVersion, L"v2.0.50727", 2) < 0 )
//							{
//								// The version is no good.
//								MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//												 "There was a problem while loading the .Net runtime:\r\n"
//												 "\tExcelClrLoadByConfig - After loading, GetCORVersion returned a version smaller than v2.0.50727.\r\n\r\n"
//												 "This problem can occur if another Excel Add-In has loaded version 1 of the runtime, or if there is a configuration file (Excel.exe.config) that specifies that an older version of the runtime should be loaded.\r\n"
//												 "Ensure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", 
//												 "ExcelDna Add-In Loader", 0);
//								hr = E_FAIL;
//							}
//							else
//							{
//								hr = S_OK;
//							}
//						}
//					}
//				}
//			}
//			FreeLibrary(hMscoree);
//		}
//	}
//	return hr;
//}


//// !!! Also does not work !!!!
//HRESULT ExcelClrLoadByTempConfigFile(ICorRuntimeHost **ppHost)
//{
//	HRESULT hr = E_FAIL;
//	char* ConfigContent =	"<configuration>"
//							"<startup>"
//							"<requiredRuntime version=\"v2.0.50727\" />"
//							"</startup>"
//							"</configuration>";
//
//    HANDLE hTempFile; 
//    wchar_t TempName[MAX_PATH];
//    wchar_t PathBuffer[MAX_PATH - 14];
//
//
//    if (GetTempPathW(MAX_PATH - 14, PathBuffer) == 0)
//	{
//		// TODO: Error
//		hr = E_FAIL;
//		DebugBreak();
//	}
//	else if (GetTempFileNameW(PathBuffer, L"DNA", 0, TempName) == 0)
//	{
//		// TODO: Error
//		DebugBreak();
//		hr = E_FAIL;
//	}
//	else
//	{
//		hTempFile = CreateFileW( TempName,
//			GENERIC_READ | GENERIC_WRITE,
//			0,
//			NULL,
//			CREATE_ALWAYS,
//			FILE_ATTRIBUTE_NORMAL,
//			NULL);
//
//		if (hTempFile == INVALID_HANDLE_VALUE) 
//		{ 
//			// TODO: Error
//			DebugBreak();
//			hr = E_FAIL;
//		} 
//		else
//		{
//			DWORD ConfigLength = (DWORD)strlen(ConfigContent);
//			DWORD BytesWritten = 0;
//			if (WriteFile(hTempFile, ConfigContent, ConfigLength, &BytesWritten, NULL) == 0)
//			{
//				CloseHandle(hTempFile);
//
//				// TODO: Error
//				DebugBreak();
//				hr = E_FAIL;
//			}
//			else
//			{
//				CloseHandle(hTempFile);
//				MessageBox(NULL, "Loading using config file.", NULL, 0);
//				MessageBoxW(NULL, TempName, NULL, 0);
//				hr = ExcelClrLoadByConfigFile(TempName, ppHost);
//				if (FAILED(hr))
//				{
//					// Message will already have been shown. Just clean up.
//					hr = E_FAIL;
//				}
//				else
//				{
//					hr = S_OK;
//				}
//			}
//
//			//DeleteFileW(TempName);
//		}
//	}
//	return hr;
//}
//
//HRESULT ExcelClrLoadByConfigFile(LPCWSTR ConfigFileName, ICorRuntimeHost** ppHost)
//{
//	HRESULT hr = E_FAIL;
//	HMODULE hMscoree = NULL;
//	hMscoree = LoadLibraryA("mscoree.dll");
//	if (hMscoree == 0)
//	{
//		MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//						 "There was a problem while loading the .Net runtime:\r\n"
//						 "\tExcelClrLoadByConfigFile - Could not load library mscoree.dll.\r\n\r\n"
//						 "This is an unexpected catastrophic failure.", 
//						 "ExcelDna Add-In Loader", 0);
//		hr = E_FAIL;					
//	}
//	else
//	{
//		pfnCorBindToCurrentRuntime CorBindToCurrentRuntime = (pfnCorBindToCurrentRuntime)GetProcAddress(hMscoree, "CorBindToCurrentRuntime");
//		if (CorBindToCurrentRuntime == 0)
//		{
//			MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//							 "There was a problem while loading the .Net runtime:\r\n"
//							 "ExcelClrLoadByConfigFile - Could not get proc address for CorBindToCurrentRuntime.\r\n\r\n"
//							 "This is an unexpected catastrophic failure.", 
//							 "ExcelDna Add-In Loader", 0);
//			hr = E_FAIL;
//		}
//		hr = CorBindToCurrentRuntime( ConfigFileName, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (LPVOID*)ppHost);
//		if (FAILED(hr))
//		{
//			MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//							 "There was a problem while loading the .Net runtime:\r\n"
//							 "ExcelClrLoadByConfigFile - CorBindToCurrentRuntime failed.\r\n\r\n"
//							 "This is an unexpected catastrophic failure.", 
//							 "ExcelDna Add-In Loader", 0);
//			hr = E_FAIL;
//		}
//		else
//		{
//			// Check the version that was loaded.
//			pfnGetCORVersion GetCORVersion = (pfnGetCORVersion)GetProcAddress(hMscoree, "GetCORVersion");
//			if (GetCORVersion == 0)
//			{
//				MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//								 "There was a problem while loading the .Net runtime:\r\n"
//								 "\tExcelClrLoadByConfigFile - Could not get proc address for GetCORVersion.\r\n"
//								 "This is an unknown, catastrophic failure.", 
//								 "ExcelDna Add-In Loader", 0);
//				hr = E_FAIL;
//			}
//			else
//			{
//				// Determine current runtime loaded, or that will be loaded:
//				WCHAR szVersion[MAX_PATH + 1];
//				DWORD dwLength = MAX_PATH;
//
//				hr = GetCORVersion(szVersion, dwLength, &dwLength);
//				if (FAILED(hr))
//				{
//					MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//									 "There was a problem while loading the .Net runtime:\r\n"
//									 "\tExcelClrLoadByConfigFile - GetCORVersion failed.\r\n"
//									 "This is an unknown, catastrophic failure.", 
//									 "ExcelDna Add-In Loader", 0);
//					hr = E_FAIL;
//				}
//				else
//				{
//#ifdef _DEBUG
//					MessageBoxW(NULL, L"Version after loading:" , L"ExcelDna Add-In Loader", 0);
//					MessageBoxW(NULL, szVersion , L"ExcelDna Add-In Loader", 0);
//#endif
//					// TODO: Proper parse of version etc.
//					if ( wcsncmp(szVersion, L"v2.0.50727", 2) < 0 )
//					{
//						// The version is no good.
//						MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//										 "There was a problem while loading the .Net runtime:\r\n"
//										 "\tExcelClrLoadByConfigFile - After loading, GetCORVersion returned a version smaller than v2.0.50727.\r\n\r\n"
//										 "This problem can occur if another Excel Add-In has loaded version 1 of the runtime, or if there is a configuration file (Excel.exe.config) that specifies that an older version of the runtime should be loaded.\r\n"
//										 "Ensure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", 
//										 "ExcelDna Add-In Loader", 0);
//						hr = E_FAIL;
//					}
//					else
//					{
//						hr = S_OK;
//					}
//				}
//			}
//		}
//		FreeLibrary(hMscoree);
//	}
//	return hr;
//}


//HRESULT ExcelClrLoadDebug(ICorRuntimeHost **ppHost)
//{
//	HRESULT hr = E_FAIL;
//	hr = CheckCORVersion();
//	
//	if (FAILED(hr))
//	{
//		hr = E_FAIL;
//	}
//	else if (hr == S_FALSE)
//	{
//	}
//	else
//	{
//		// Load COR using default loading mechanism
//		hr = ExcelClrLoad(ppHost);
//	}
//
//
//	if (hr = S_OK)
//	{
//		// Check version
//	}
//
//	//// Attempt to load a runtime that is compatible with the release version of .Net 2.0.
//	//hr = CorBindToRuntimeEx(L"v2.0.50727", L"wks", NULL, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (LPVOID*)ppHost);
//	////if (hr == CLR_E_SHIM_RUNTIMELOAD)
//	////{
//	////	MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while loading the .Net runtime:\r\n\tExcelClrLoader - CorBindToRuntimeEx could not load a version compatible with v2.0.50727.\r\nThis error can occur if the .Net runtime is not installed, if version 2 or later is not installed, if the AppPatch registry entries are incorrect, or if there is a config file for Excel retricting the versions of the .Net runtime that can be loaded.\r\n\r\nEnsure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", "ExcelDna Add-In Loader", 0);
//	////	return E_FAIL;
//	////}
//	//if (FAILED(hr))
//	//{
//	//	MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while loading the .Net runtime:\r\n\tExcelClrLoader - CorBindToRuntimeEx failed.\r\nThis error can occur if the .Net runtime is not installed, if version 2 or later is not installed, if the AppPatch registry entries are incorrect, or if there is a config file for Excel retricting the versions of the .Net runtime that can be loaded.\r\n\r\nEnsure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", "ExcelDna Add-In Loader", 0);
//	//	return E_FAIL;
//	//}
//		hMscoree = LoadLibraryA("mscoree.dll");
//		if (hMscoree == 0)
//		{
//			MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while loading the .Net runtime:\r\n\tExcelClrLoader - Could not load library mscoree.dll.\r\nThis error can occur if the .Net runtime is not installed.\r\n\r\nEnsure that the .Net runtime version 2 or later is installed before loading the add-in.", "ExcelDna Add-In Loader", 0);
//			return E_FAIL;
//		}
//		pfnCorBindToRuntimeEx CorBindToRuntimeEx = (pfnCorBindToRuntimeEx)GetProcAddress(hMscoree, "CorBindToRuntimeEx");
//		if (CorBindToRuntimeEx == 0)
//		{
//			MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while loading the .Net runtime:\r\n\tExcelClrLoader - Could not get proc address for CorBindToRuntimeEx.\r\nThis error can occur if the .Net runtime is not installed, or if version 2 or later is not installed.\r\n\r\nEnsure that the .Net runtime version 2 or later is installed before loading the add-in.", "ExcelDna Add-In Loader", 0);
//			hr E_FAIL;
//		}
//		else
//		{
//
//	hr = ExcelClrLoadByConfig(ppHost);
//
//	// Check the version
//	WCHAR szVersion[MAX_PATH + 1];
//	DWORD dwLength = MAX_PATH;
//	hr = GetCORVersion(szVersion, dwLength, &dwLength);
//	if (FAILED(hr))
//	{
//		MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while loading the .Net runtime:\r\n\tExcelClrLoader - GetCORVersion failed.\r\nThis is an unknown, catastrophic failure.", "ExcelDna Add-In Loader", 0);
//		return E_FAIL;
//	}
//
//	MessageBoxW(NULL, L"Version after loading by CorBindToRuntimeEx(v2.0.50727):" , L"ExcelDna Add-In Loader", 0);
//	MessageBoxW(NULL, szVersion , L"ExcelDna Add-In Loader", 0);
//
//	MessageBoxW(NULL, L"Version after loading by CorBindToRuntimeByCfg(v2.0.50727):" , L"ExcelDna Add-In Loader", 0);
//	MessageBoxW(NULL, szVersion , L"ExcelDna Add-In Loader", 0);
//
//		// Compare the version with our requirement
//	// TODO: Do this right -- will not work for v13.8.12345
//	if ( wcsncmp(szVersion, L"v2.0.50727", 2) < 0 )
//	{
//		// The version is no good.
//		MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while loading the .Net runtime:\r\n\tExcelClrLoader - GetCORVersion returned a version smaller than v2.0.50727.\r\nThis error can occur if the .Net runtime is not installed, if version 2 or later is not installed, if the AppPatch registry entries are incorrect, or if there is a config file for Excel retricting the versions of the .Net runtime that can be loaded.\r\n\r\nEnsure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", "ExcelDna Add-In Loader", 0);
//		return E_FAIL;
//	}
//
//	// Now the right version might be loaded, or some other version.
//	hr = (*ppHost)->Start();
//	if (FAILED(hr))
//	{
//		MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while starting the .Net runtime:\r\n\tExcelClrLoader - Host->Start failed.\r\nThis is an unknown, catastrophic failure.", "ExcelDna Add-In Loader", 0);
//		return E_FAIL;
//	}
//	return S_OK;
//
//}
//


// returns E_FAIL if there was a fatal error - messages already shown.
// returns S_OK if the version reported by GetCORVersion is v2.0.50727 or later, and a current runtime was loaded.
// returns S_FALSE if the version reported by GetCORVersion is older, fixups might be needed.
HRESULT ExcelClrLoadv2(ICorRuntimeHost **ppHost)
{
	HRESULT hr = E_FAIL;
	HMODULE hMscoree = NULL;

	hMscoree = LoadLibrary(L"mscoree.dll");
	if (hMscoree == 0)
	{
		MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime.\r\n"
						 L"There was a problem while loading the .Net runtime:\r\n"
						 L"\tExcelClrLoader - Could not load library mscoree.dll.\r\n"
						 L"This error can occur if the .Net runtime is not installed.\r\n\r\n"
						 L"Ensure that the .Net runtime version 2.0 or later is installed before loading the add-in.", 
						 L"ExcelDna Add-In Loader", 0);
		hr = E_FAIL;
	}
	else
	{
		pfnGetCORVersion GetCORVersion = (pfnGetCORVersion)GetProcAddress(hMscoree, "GetCORVersion");
		if (GetCORVersion == 0)
		{
			MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime.\r\n"
							 L"There was a problem while loading the .Net runtime:\r\n"
							 L"\tExcelClrLoader - Could not get procedure address for GetCORVersion.\r\n"
							 L"This is an unexpected, catastrophic failure.", 
							 L"ExcelDna Add-In Loader", 0);
			hr = E_FAIL;
		}
		else
		{
			// Display current runtime loaded, or that will be loaded:
			WCHAR szVersion[MAX_PATH + 1];
			DWORD dwLength = MAX_PATH;

//			hr = GetCORVersion(szVersion, dwLength, &dwLength);
//			if (hr == CLR_E_SHIM_RUNTIMELOAD)
//			{
//				// Might have a bad version specified in the excel.exe.config file,
//				MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime.\r\n"
//								 L"There was a problem while loading the .Net runtime:\r\n"
//								 L"\tExcelClrLoader - GetCORVersion failed.\r\n\r\n"
//								 L"This can occur if there is an incorrect configuration file (Excel.exe.config) or if there is a problem with the .Net installation.\r\n\r\n"
//								 L"The loader will attempt to force loading of .Net version 2.0.50727 ...",
//								 L"ExcelDna Add-In Loader", 0);
//				hr = S_FALSE;
//			}
//			else if (FAILED(hr))
//			{
//				MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime.\r\n"
//								 L"There was a problem while loading the .Net runtime:\r\n"
//								 L"\tExcelClrLoader - GetCORVersion failed.\r\n\r\n"
//								 L"This can occur if there is an incorrect configuration file (Excel.exe.config) or if there is a problem with the .Net installation.\r\n"
//								 L"Ensure that the .Net runtime version 2 or later is installed, and will be loaded correctly under Excel, before loading the add-in.", 
//								 L"ExcelDna Add-In Loader", 0);
//				hr = E_FAIL;
//			}
//			else
//			{
//#ifdef _DEBUG
//				MessageBoxW(NULL, L"Version before loading:" , L"ExcelDna Add-In Loader", 0);
//				MessageBoxW(NULL, szVersion , L"ExcelDna Add-In Loader", 0);
//#endif
//				// TODO: Proper parse of version etc.
//				if ( wcsncmp(szVersion, L"v2.0.50727", 2) < 0 )
//				{
//					//// The version is no good.
//					//MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
//					//				 "There was a problem while loading the .Net runtime:\r\n"
//					//				 "\tExcelClrLoader - GetCORVersion returned a version smaller than v2.0.50727.\r\nThis error can occur if the .Net runtime is not installed, if version 2 or later is not installed, if the AppPatch registry entries are incorrect, or if there is a config file for Excel retricting the versions of the .Net runtime that can be loaded.\r\n\r\nEnsure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", "ExcelDna Add-In Loader", 0);
//					hr = S_FALSE;
//				}
//				else
//				{
					// Load the runtime
					pfnCorBindToRuntimeEx CorBindToRuntimeEx = (pfnCorBindToRuntimeEx)GetProcAddress(hMscoree, "CorBindToRuntimeEx");
					if (CorBindToRuntimeEx == 0)
					{
						// Should not happen - we have already checked that .Net 2.0 is installed
						MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
										 L"There was a problem while loading the .Net runtime:\r\n"
										 L"ExcelClrLoader - Could not get proc address for CorBindToRuntimeEx.\r\n\r\n"
										 L"This is an unexpected catastrophic failure.", 
										 L"ExcelDna Add-In Loader", 0);
						hr = E_FAIL;
					}
					else
					{
						// Attempt to load a runtime that is compatible with the release version of .Net 2.0.
						hr = CorBindToRuntimeEx(L"v2.0.50727", L"wks", NULL, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (LPVOID*)ppHost);
						//if (hr == CLR_E_SHIM_RUNTIMELOAD)
						//{
						//	MessageBox(NULL, "This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\nThere was a problem while loading the .Net runtime:\r\n\tExcelClrLoader - CorBindToRuntimeEx could not load a version compatible with v2.0.50727.\r\nThis error can occur if the .Net runtime is not installed, if version 2 or later is not installed, if the AppPatch registry entries are incorrect, or if there is a config file for Excel retricting the versions of the .Net runtime that can be loaded.\r\n\r\nEnsure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", "ExcelDna Add-In Loader", 0);
						//	return E_FAIL;
						//}
						if (FAILED(hr))
						{
							MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime.\r\n"
											 L"There was a problem while loading the .Net runtime:\r\n"
											 L"\tExcelClrLoader - CorBindToRuntimeEx failed.\r\n"
											 L"This error can occur if the .Net runtime is not installed, if version 2 or later is not installed, if the AppPatch registry entries are incorrect, or if there is a config file for Excel retricting the versions of the .Net runtime that can be loaded.\r\n\r\n"
											 L"Ensure that the .Net runtime version 2 or later is installed, and will be loaded correctly, before loading the add-in.", 
											 L"ExcelDna Add-In Loader", 0);
							hr = E_FAIL;
						}
						else
						{
							// Check the version that is now loaded ...
							hr = GetCORVersion(szVersion, dwLength, &dwLength);
							if (FAILED(hr))
							{
								MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime.\r\n"
												 L"There was a problem while loading the .Net runtime:\r\n"
												 L"\tExcelClrLoader - GetCORVersion failed.\r\n"
												 L"This is an unknown, catastrophic failure.", 
												 L"ExcelDna Add-In Loader", 0);
								hr = E_FAIL;
							}
							else
							{
#ifdef _DEBUG
//								MessageBoxW(NULL, L"Version after loading:" , L"ExcelDna Add-In Loader", 0);
//								MessageBoxW(NULL, szVersion , L"ExcelDna Add-In Loader", 0);
#endif
								// TODO: Proper parse of version etc.
								if ( DetectFxReadMajorVersion(szVersion) < 2 )
								{
									// The version is no good.
									MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
													 L"Although the required runtime is installed, it could not be loaded successfully.\r\n"
													 L"This is possible if a configuration file or another managed add-in forces an older version of the runtime to load (only one version of the runtime can be loaded into the process).\r\n\r\n"
													 L"Review the other add-ins that are loaded, or ensure that the .Net 2.0 runtime loads by setting an approriate 'supportedRuntime' entry in a configuration file (Excel.exe.config).\r\n\r\n"
													 L"You will need to restart Excel to load the correct version of the runtime.",
													L"ExcelDna Add-In Loader", 0);
									hr = E_FAIL;
								}
								else
								{
									hr = S_OK;
								}
							}
//						}
//					}
				}
			}
		}
		FreeLibrary(hMscoree);
	}
	return hr;
}

// ExcelClrLoad returns S_FALSE if a v2+ CLR is already loaded, 
// S_OK if the CLR was loaded successfully, E_FAIL if a CLR could not be loaded.
// ExcelClrLoad shows diagnostic MessageBoxes if needed.
HRESULT ExcelClrLoad(ICorRuntimeHost **ppHost)
{

//	HRESULT hrr = RepairRegistryAppPatch();

	HRESULT hr = E_FAIL;
	// Try to load the v2 runtime
	hr = ExcelClrLoadv2(ppHost);
	if (hr == E_FAIL)
	{
		// Bad error, could not load - message already shown.
	}
	else if (hr == S_FALSE)
	{
		// Could not load the right version
		// Check whether version 2 is installed
		if (!DetectFxIsNet20Installed())
		{
			MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
							 L"The required runtime was not detected on this machine.\r\n\r\n"
							 L"Ensure that the .Net runtime version 2.0 or later is correctly installed before loading the Add-In.", 
							 L"ExcelDna Add-In Loader", 0);
			hr = E_FAIL;
		}
		else
		{
#ifdef _DEBUG
			MessageBox(NULL, L"Net20Installed!", NULL, 0);
#endif
			// Check whether a version is already running
			if (GetModuleHandle(L"mscorwks") != NULL)
			{
				MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
								 L"Although the required runtime is installed, an older version has been loaded by Excel (only one version of the runtime can be loaded).\r\n"
								 L"This is possible if another managed add-in caused the runtime to load, using the incorrectly installed AppPatch policy.\r\n\r\n"
								 L"Check the other add-ins that are loaded, or ensure that the .Net 2.0 runtime loads by setting an appropriate 'supportedRuntime' entry in a configuration file (Excel.exe.config).",
								 L"ExcelDna Add-In Loader", 0);
				hr = E_FAIL;
			}
			else
			{
				// There is still hope . . .
				hr = RepairRegistryAppPatch();
				if (SUCCEEDED(hr))
				{
					// Try to load 2.0 again
					hr = ExcelClrLoadv2(ppHost);
					if (hr == S_FALSE)
					{
						// Could not load even after registry repair - give up.
						MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
										 L"Although the required runtime is installed, it cannot be loaded successfully (even after successfully repairing the AppPatch entry in the registry).\r\n"
										 L"This is possible if a configuration file or another managed add-in forces an older version of the runtime to load (only one version of the runtime can be loaded).\r\n\r\n"
										 L"Check the other add-ins that are loaded, or ensure that the .Net 2.0 runtime loads by setting an approriate 'supportedRuntime' entry in a configuration file (Excel.exe.config).",
										 L"ExcelDna Add-In Loader", 0);
						hr = E_FAIL;
					}
					// otherwise succeeded or failed.
				}
				else
				{
					// Could not repair registry
					// Try to write config file to the excel directory
					// Check whether Excel.exe.config file exists.

					HANDLE hConfigFile = CreateFile(L"excel.exe.config", FILE_READ_DATA, 0, NULL, OPEN_EXISTING, 0, NULL);
					if (hConfigFile != INVALID_HANDLE_VALUE)
					{
						MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
										 L"Although the required runtime is installed, it cannot be loaded successfully.\r\n"
										 L"This is possible if the configuration file (Excel.exe.config) that is present or another managed add-in forces an older version of the runtime to load (only one version of the runtime can be loaded).\r\n\r\n"
										 L"Check the other add-ins that are loaded, or ensure that the .Net 2.0 runtime loads by setting an approriate 'supportedRuntime' entry in the configuration file (Excel.exe.config).",
										 L"ExcelDna Add-In Loader", 0);
						hr = E_FAIL;
					}
					else
					{
						CloseHandle(hConfigFile);
					}
				}
			}
		}
		
	}

	// If all is fine now, also start the CLR
	if (hr == S_OK)
	{
 		hr = (*ppHost)->Start();
		if (FAILED(hr))
		{
			MessageBox(NULL, L"This Excel Add-In requires the .Net 2.0 runtime to be loaded.\r\n"
							 L"The .Net runtime could not be started:\r\n"
							 L"\tExcelClrLoad - Host->Start failed.\r\n\r\n"
							 L"This is an unexpected, catastrophic failure.", 
							 L"ExcelDna Add-In Loader", 0);
			hr = E_FAIL;
		}
		else
		{
			hr = S_OK;
		}
	}
	return hr;
}

HRESULT RepairRegistryAppPatch()
{
	HRESULT hr = E_FAIL;
	HKEY hkAppPatchExcel;

	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\.NETFramework\\policy\\AppPatch\\v2.0.50727.00000\\excel.exe\\{2CCAA9FE-6884-4AF2-99DD-5217B94115DF}", 
					 0, KEY_READ | KEY_WRITE, &hkAppPatchExcel) != ERROR_SUCCESS)
	{
		// Could not open the registry (with required access) - maybe security....
		hr = E_FAIL;
	}
	else
	{
		DWORD BufferLength;
		BYTE  Buffer[1000];
		DWORD Type;

		BufferLength = 1000;
		if (RegQueryValueEx(hkAppPatchExcel, L"Minimum File Version Number", NULL, &Type, Buffer, &BufferLength) != ERROR_SUCCESS)
		{
			hr = E_FAIL;
		}
		else
		{
			if (RegSetValueEx(hkAppPatchExcel, L"Minimum File Version", NULL, REG_SZ, Buffer, BufferLength) != ERROR_SUCCESS)
			{
				hr = E_FAIL;
			}
			else
			{
				// Do Maximum
				BufferLength = 1000;
				if (RegQueryValueEx(hkAppPatchExcel, L"Maximum File Version Number", NULL, &Type, Buffer, &BufferLength) != ERROR_SUCCESS)
				{
					hr = E_FAIL;
				}
				else
				{
					if (RegSetValueEx(hkAppPatchExcel, L"Maximum File Version", NULL, REG_SZ, Buffer, BufferLength) != ERROR_SUCCESS)
					{
						hr = E_FAIL;
					}
					else
					{
						hr = S_OK;
					}
				}
			}
		}

		RegCloseKey(hkAppPatchExcel);
	}
	return hr;
}