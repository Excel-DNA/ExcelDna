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
#include "LoadFx.h"

typedef HRESULT (STDAPICALLTYPE *pfnCorBindToRuntimeEx)(        
									LPWSTR pwszVersion,   
									LPWSTR pwszBuildFlavor, 
									DWORD flags,            
									REFCLSID rclsid,      
									REFIID riid,    
									LPVOID* ppv );

typedef HRESULT (STDAPICALLTYPE *pfnGetCORVersion)(
										LPWSTR szBuffer, 
                                        DWORD cchBuffer,
                                        DWORD* dwLength);

HRESULT LoadClrv2(ICorRuntimeHost **ppHost);


// returns E_FAIL if there was a fatal error - messages already shown.
// returns S_OK if the version reported by GetCORVersion is v2.0.50727 or later, and a current runtime was loaded.
// returns S_FALSE if the version reported by GetCORVersion is older, fixups might be needed.
HRESULT LoadClrv2(ICorRuntimeHost **ppHost)
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
				}
			}
		}
		FreeLibrary(hMscoree);
	}
	return hr;
}

// LoadClr returns S_FALSE if a v2+ CLR is already loaded, 
// S_OK if the CLR was loaded successfully, E_FAIL if a CLR could not be loaded.
// LoadClr shows diagnostic MessageBoxes if needed.
HRESULT LoadClr(ICorRuntimeHost **ppHost)
{

	HRESULT hr = E_FAIL;
	// Try to load the v2 runtime
	hr = LoadClrv2(ppHost);
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
