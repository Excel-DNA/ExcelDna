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
#include "ExcelDna.h"
#include "ExcelDnaLoader.h"
#include "resource.h"

#define CountOf(x) sizeof(x)/sizeof(*x)
#define MAX_MSG = 1024;

static HMODULE hModuleCurrent;
// These don't use ATL classes to give us explicit control over when CLR is called
static IUnknown* pAppDomain_ForUnload = NULL;
static ICorRuntimeHost* pHost_ForUnload = NULL;
HRESULT LoadClr20(ICorRuntimeHost **ppHost);
void ShowMessage(int headerId, int bodyId, int footerId, HRESULT hr = S_OK);
CString AddInFullPath();

typedef HRESULT (STDAPICALLTYPE *pfnGetCORVersion)(LPWSTR pBuffer, 
                                         DWORD cchBuffer,
                                         DWORD* dwLength);

typedef HRESULT (STDAPICALLTYPE *pfnGetVersionFromProcess)(
										 HANDLE hProcess,
										 LPWSTR pBuffer, 
                                         DWORD cchBuffer,
                                         DWORD* dwLength);

typedef HRESULT (STDAPICALLTYPE *pfnCorBindToRuntimeEx)(        
									LPWSTR pwszVersion,   
									LPWSTR pwszBuildFlavor, 
									DWORD flags,            
									REFCLSID rclsid,      
									REFIID riid,    
									LPVOID* ppv );

bool XlLibraryInitialize(XlAddInExportInfo* pExportInfo)
{
	HRESULT hr;
	CComPtr<ICorRuntimeHost> pHost;

	hr = LoadClr20(&pHost);
	if (FAILED(hr) || pHost == NULL)
	{
		// LoadClr20 shows diagnostic MessageBoxes if needed.
		// Perhaps remember that we are not loaded?
		return 0;
	}

	// If all is fine now, also start the CLR (always safe to do again.
	hr = pHost->Start();
	if (FAILED(hr))
	{
		ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
					IDS_MSG_BODY_HOSTSTART,
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	CString addInFullPath = AddInFullPath();

	CPath xllDirectory(addInFullPath);
	xllDirectory.RemoveFileSpec();

	CComPtr<IUnknown> pAppDomainSetupUnk;
	hr = pHost->CreateDomainSetup(&pAppDomainSetupUnk);
	if (FAILED(hr) || pAppDomainSetupUnk == NULL)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_APPDOMAINSETUP, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	CComQIPtr<IAppDomainSetup> pAppDomainSetup = pAppDomainSetupUnk;

	hr = pAppDomainSetup->put_ApplicationBase(CComBSTR(xllDirectory));
	if (FAILED(hr))
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_APPLICATIONBASE, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}
	
	CComBSTR configFileName = addInFullPath;
	configFileName.Append(L".config");
	pAppDomainSetup->put_ConfigurationFile(configFileName);

	CComBSTR appDomainName = L"ExcelDna: ";
	appDomainName.Append(addInFullPath);
	pAppDomainSetup->put_ApplicationName(appDomainName);

	IUnknown *pAppDomainUnk = NULL;
	hr = pHost->CreateDomainEx(appDomainName, pAppDomainSetupUnk, 0, &pAppDomainUnk);
	if (FAILED(hr) || pAppDomainUnk == NULL)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_APPDOMAIN, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	CComQIPtr<_AppDomain> pAppDomain(pAppDomainUnk);

	// Load plan for ExcelDna.Loader:
	// Try AppDomain.Load with the name ExcelDna.Loader.
	// Then if it does not work, we will try to load from a known resource in the .xll.

	CComPtr<_Assembly> pExcelDnaLoaderAssembly;
	hr = pAppDomain->Load_2(CComBSTR(L"ExcelDna.Loader"), &pExcelDnaLoaderAssembly);
	if (FAILED(hr) || pExcelDnaLoaderAssembly == NULL)
	{
		HRSRC hResInfoLoader = FindResource(hModuleCurrent, L"EXCELDNA_LOADER", L"ASSEMBLY");
		if (hResInfoLoader == NULL)
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_MISSINGEXCELDNALOADER, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return 0;
		}
		HGLOBAL hLoader = LoadResource(hModuleCurrent, hResInfoLoader);
		void* pLoader = LockResource(hLoader);
		ULONG sizeLoader = (ULONG)SizeofResource(hModuleCurrent, hResInfoLoader);
		
		CComSafeArray<BYTE> bytesLoader;
		bytesLoader.Add(sizeLoader, (byte*)pLoader);

		hr = pAppDomain->Load_3(bytesLoader, &pExcelDnaLoaderAssembly);
		if (FAILED(hr))
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_EXCELDNALOADER, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return 0;
		}

		CComBSTR pFullName;
		hr = pExcelDnaLoaderAssembly->get_FullName(&pFullName);
		if (FAILED(hr))
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_EXCELDNALOADERNAME, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return 0;
		}
	}
	
	CComPtr<_Type> pXlAddInType;
	hr = pExcelDnaLoaderAssembly->GetType_2(CComBSTR(L"ExcelDna.Loader.XlAddIn"), &pXlAddInType);
	if (FAILED(hr) || pXlAddInType == NULL)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_XLADDIN, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	CComSafeArray<VARIANT> initArgs;
	initArgs.Add(CComVariant((INT32)pExportInfo));
	initArgs.Add(CComVariant((INT32)hModuleCurrent));
	initArgs.Add(CComVariant(addInFullPath.AllocSysString()));
	CComVariant initRetVal;
	CComVariant target;
	hr = pXlAddInType->InvokeMember_3(CComBSTR("Initialize"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, target, initArgs, &initRetVal);
	if (FAILED(hr))
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_XLADDININIT, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	pHost_ForUnload = pHost.Detach();
	pAppDomain_ForUnload = (IUnknown*)pAppDomain.Detach();

	return initRetVal.boolVal == 0 ? false : true;
}

void XlLibraryUnload()
{
	if (pHost_ForUnload != NULL)
	{
		HRESULT hr = pHost_ForUnload->UnloadDomain(pAppDomain_ForUnload);
		pAppDomain_ForUnload->Release();
		pAppDomain_ForUnload = NULL;
		pHost_ForUnload->Release();
		pHost_ForUnload = NULL;
		if (FAILED(hr))
		{
#if _DEBUG
			DebugBreak();
#endif
		}
	}
}

HRESULT LoadClr20(ICorRuntimeHost **ppHost)
{
	HRESULT hr = E_FAIL;
	HMODULE hMscoree = NULL;

	hMscoree = LoadLibrary(L"mscoree.dll");
	if (hMscoree == 0)
	{
		ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
					IDS_MSG_BODY_LOADMSCOREE, 
					IDS_MSG_FOOTER_ENSURECLR20 );
		hr = E_FAIL;
	}
	else
	{
		// Load the runtime
		pfnCorBindToRuntimeEx CorBindToRuntimeEx = (pfnCorBindToRuntimeEx)GetProcAddress(hMscoree, "CorBindToRuntimeEx");
		if (CorBindToRuntimeEx == 0)
		{
			ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
						IDS_MSG_BODY_NOCORBIND, 
						IDS_MSG_FOOTER_UNEXPECTED );
			hr = E_FAIL;
		}
		else
		{
			// Attempt to load a runtime that is compatible with the release version of .Net 2.0.
			hr = CorBindToRuntimeEx(L"v2.0.50727", L"wks", NULL, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (LPVOID*)ppHost);
			if (FAILED(hr))
			{
				// Could not load the right version
				// Check whether version 2 is installed
				if (!DetectFxIsNet20Installed())
				{
					ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
								IDS_MSG_BODY_NONET20,
								IDS_MSG_FOOTER_ENSURECLR20);
					hr = E_FAIL;
				}
				else
				{
					// Check whether a version is already running
					if (GetModuleHandle(L"mscorwks") != NULL)
					{
						ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
									IDS_MSG_BODY_OLDVERSION,
									IDS_MSG_FOOTER_OLDVERSION);
						hr = E_FAIL;
					}
					else
					{
						ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
							IDS_MSG_BODY_CORBINDFAILED, 
							IDS_MSG_FOOTER_ENSURECLR20ANDLOAD, 
							hr);						
						//// Unknown load failure
						//ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
						//			IDS_MSG_BODY_UNKNOWNLOADFAIL,
						//			IDS_MSG_FOOTER_UNEXPECTED);
						hr = E_FAIL;
					}
				}
				//hr = E_FAIL;
			}
			else
			{
				// Check the version that is now loaded ...
				pfnGetCORVersion GetCORVersion = (pfnGetCORVersion)GetProcAddress(hMscoree, "GetCORVersion");
				if (GetCORVersion == 0)
				{
					ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
								IDS_MSG_BODY_NOCORVERSION, 
								IDS_MSG_FOOTER_UNEXPECTED );
					hr = E_FAIL;
				}
				else
				{
					// Display current runtime loaded
					WCHAR szVersion[MAX_PATH + 1];
					DWORD dwLength = MAX_PATH;
					hr = GetCORVersion(szVersion, dwLength, &dwLength);
					if (FAILED(hr))
					{
						ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
									IDS_MSG_BODY_CORVERSIONFAILED, 
									IDS_MSG_FOOTER_UNEXPECTED,
									hr);
						hr = E_FAIL;
					}
					else
					{
						if ( DetectFxReadMajorVersion(szVersion) < 2 )
						{
							// The version is no good.
							ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
										IDS_MSG_BODY_WRONGVERSIONLOADED, 
										IDS_MSG_FOOTER_REVIEWADDINS);
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

void SetCurrentModule(HMODULE hModule)
{
	hModuleCurrent = hModule;
}

struct FindExcelWindowParam
{
	DWORD processId;
	HWND  hwndFound;
};

BOOL WINAPI FindExcelWindowCallback(HWND hwnd, LPARAM lParam)
{
	FindExcelWindowParam* pParam = (FindExcelWindowParam*)lParam;
	DWORD processId = 0;
	GetWindowThreadProcessId(hwnd, &processId);
	if (processId == pParam->processId)
	{
		CString className;
		LPTSTR pBuffer = className.GetBuffer(10);
		DWORD count = RealGetWindowClass(hwnd, pBuffer, 10);
		className.ReleaseBuffer(count);
		if (className == L"XLMAIN")
		{
			pParam->hwndFound = hwnd;
			SetLastError(0);
			return FALSE;
		}
	}
	return TRUE;
}

HWND FindCurrentExcelWindow()
{
	FindExcelWindowParam param;
	param.processId = GetCurrentProcessId();
	param.hwndFound = NULL;

	EnumWindows(FindExcelWindowCallback, (LPARAM)&param);
	return param.hwndFound;
}

void ShowMessageError(HWND hwndParent)
{
	MessageBox(hwndParent, L"There was problem while loading the add-in. \r\nA detailed message could not be displayed.", L"Add-In Loader", MB_ICONEXCLAMATION);
}

void ShowMessage(int headerId, int bodyId, int footerId, HRESULT hr)
{
	HWND hwndExcel = FindCurrentExcelWindow();
	try
	{
		CString addInFullPath = AddInFullPath();

		CPath addInFileName = addInFullPath;
		addInFileName.StripPath();

		CString msgTitle;
		msgTitle.FormatMessage(IDS_MSG_TITLE, addInFileName);

		CString header;
		header.LoadString(headerId);
		CString body;
		body.LoadString(bodyId);
		CString footer;
		footer.LoadString(footerId);
		CString hresult = "";
		if (hr != S_OK)
		{
			hresult.FormatMessage(IDS_MSG_HRESULT, hr);
		}

		CString msg;
		msg.FormatMessage(IDS_MSG_TEMPLATE, header, body, footer, hresult, addInFullPath);

		MessageBox(hwndExcel, msg, msgTitle, MB_ICONEXCLAMATION);
	}
	catch (...)
	{
		ShowMessageError(hwndExcel);
	}
}

CString AddInFullPath()
{
	CString addInFullPath;
	LPTSTR pBuffer = addInFullPath.GetBuffer(MAX_PATH);
	DWORD count = GetModuleFileName(hModuleCurrent, pBuffer, MAX_PATH);
	addInFullPath.ReleaseBuffer(count); // pBuffer is now invalid
	return addInFullPath;
}