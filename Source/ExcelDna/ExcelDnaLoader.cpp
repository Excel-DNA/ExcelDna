/*
  Copyright (C) 2005-2010 Govert van Drimmelen

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


// TODO: Change to the .Net 2.0+ hosting interface IClrRuntimeHost, 
// so that we can set the safe AppDomain flags when loading.

#include "stdafx.h"
#include "DetectFx.h"
#include "ExcelDna.h"
#include "ExcelDnaLoader.h"
#include "resource.h"

#define CountOf(x) sizeof(x)/sizeof(*x)
const int MAX_MSG = 1024;
const CString CLR_VERSION_20 = L"v2.0.50727";

static HMODULE hModuleCurrent;
// These don't use ATL classes to give us explicit control over when CLR is called
static IUnknown* pAppDomain_ForUnload = NULL;
static ICorRuntimeHost* pHost_ForUnload = NULL;
// Temp file to be used if we need to write .config from resources.
static CString tempConfigFileName = "";

// Forward declarations for functions defined in this file.
HRESULT LoadClr(CString clrVersion, ICorRuntimeHost **ppHost);
HRESULT LoadClrMeta(CString clrVersion, ICLRMetaHost* pMetaHost, ICorRuntimeHost **ppHost);
HRESULT LoadClr20(ICorRuntimeHost **ppHost);
void ShowMessage(int headerId, int bodyId, int footerId, HRESULT hr = S_OK);
CString AddInFullPath();
HRESULT CreateTempFile(void* pBuffer, DWORD nBufSize, CString& fileName);
HRESULT DeleteTempFile(CString fileName);

HRESULT ReadClrOptions(CString& clrVersion, bool& shadowCopyFiles);
HRESULT ReadDnaHeader(CString& header);
HRESULT ParseDnaHeader(CString header, CString& runtimeVersion, bool& shadowCopyFiles);
HRESULT ReadAttributeValue(CString tag, CString attributeName, CString& attributeValue);

BOOL IsBufferUTF8(BYTE* buffer, DWORD bufferLength);
CStringW UTF8toUTF16(const CStringA& utf8);

// COR function pointer typedefs.
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

typedef HRESULT (STDAPICALLTYPE *pfnCLRCreateInstance)(        
									REFCLSID  clsid,
									REFIID riid,
									LPVOID* ppInterface );

// Ensure the CLR is loaded, create a new AppDomain, get the manager loader running,
// and do the ExportInfo hook-up.
bool XlLibraryInitialize(XlAddInExportInfo* pExportInfo)
{
	HRESULT hr;
	CComPtr<ICorRuntimeHost> pHost;
	CString clrVersion;
	bool shadowCopyFiles;
	
	hr = ReadClrOptions(clrVersion, shadowCopyFiles);
	if (FAILED(hr))
	{
		// SelectClrVersion shows diagnostic MessageBoxes if needed.
		// Perhaps remember that we are not loaded?
		return 0;
	}
	hr = LoadClr(clrVersion, &pHost);
	if (FAILED(hr) || pHost == NULL)
	{
		// LoadClr shows diagnostic MessageBoxes if needed.
		// Perhaps remember that we are not loaded?
		return 0;
	}

	// If all is fine now, also start the CLR (always safe to do again).
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

	// Create and populate AppDomainSetup
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

	hr = pAppDomainSetup->put_ShadowCopyFiles(CComBSTR(shadowCopyFiles ? L"true" : L"false"));
	if (FAILED(hr))
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_SHADOWCOPYFILES, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	// AppDomainSetup.ApplicationName = "ExcelDna: c:\MyAddins\MyAddIn.xll";
	CComBSTR appDomainName = L"ExcelDna: ";
	appDomainName.Append(addInFullPath);
	pAppDomainSetup->put_ApplicationName(appDomainName);


	// Check if a .config file exists next to the .xll as MyAddIn.xll.config. Use it if it exists.
	CComBSTR configFileName = addInFullPath;
	configFileName.Append(L".config");
	if (ATLPath::FileExists(configFileName))
	{
		pAppDomainSetup->put_ConfigurationFile(configFileName);
	}
	else
	{
		// Try to load .config file from resources, store into a temp file
		HRSRC hResConfig = FindResource(hModuleCurrent, L"__MAIN__", L"CONFIG");
		if (hResConfig != NULL)
		{
			HGLOBAL hConfig = LoadResource(hModuleCurrent, hResConfig);
			void* pConfig = LockResource(hConfig);
			DWORD sizeConfig = SizeofResource(hModuleCurrent, hResConfig);

			hr = CreateTempFile(pConfig, sizeConfig, tempConfigFileName);
			if (SUCCEEDED(hr))
			{
				pAppDomainSetup->put_ConfigurationFile( CComBSTR(tempConfigFileName) );
			}
			// tempConfigFile will be deleted after the AppDomain has been unloaded.
		}
		else
		{
			// No config file - no problem.
		}
	}

	// Maybe for .NET 4 we need to make a sandboxed AppDomain.
	// I think this is only possible from within managed code.
	// So my plan is to make an additional loader AppDomain, that will create the sandboxed AppDomain,
	// and return the AppDomainId. 
	// We then call into the sandboxed AppDomain from the managed code.
	// Still need to be able to unload from here....?

	// Some ideas: http://www.mmowned.com/forums/world-of-warcraft/bots-programs/memory-editing/300096-net-managed-assembly-removal.html

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
	CComSafeArray<BYTE> bytesLoader;
	bool loadLoaderFromBytes = false;

	hr = pAppDomain->Load_2(CComBSTR(L"ExcelDna.Loader"), &pExcelDnaLoaderAssembly);
	if (FAILED(hr) || pExcelDnaLoaderAssembly == NULL)
	{
		HRSRC hResInfoLoader = FindResource(hModuleCurrent, L"EXCELDNA.LOADER", L"ASSEMBLY");
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
		loadLoaderFromBytes = true;

		// Is this just for debugging?
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
	
	CComPtr<_Type> pAppDomainHelperType;
	hr = pExcelDnaLoaderAssembly->GetType_2(CComBSTR(L"ExcelDna.Loader.AppDomainHelper"), &pAppDomainHelperType);
	if (FAILED(hr) || pAppDomainHelperType == NULL)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_XLADDIN, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	CComSafeArray<VARIANT> sbArgs;
	CComVariant sbRetVal;
	CComVariant sbTarget;
	hr = pAppDomainHelperType->InvokeMember_3(CComBSTR("CreateFullTrustSandbox"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, sbTarget, sbArgs, &sbRetVal);
	if (FAILED(hr))
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_XLADDININIT, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	CComQIPtr<_AppDomain> pSandbox(sbRetVal.punkVal);
	CComBSTR sandboxName;
	pSandbox->get_FriendlyName(&sandboxName);

	// Now load the Loader assembly in the Sandboxed AppDomain, and call XlAddIn.Initialize.
	CComPtr<_Assembly> pExcelDnaLoaderAssemblyInSandbox;
	if (!loadLoaderFromBytes)
	{
		hr = pSandbox->Load_2(CComBSTR(L"ExcelDna.Loader"), &pExcelDnaLoaderAssemblyInSandbox);
		if (FAILED(hr))
		{
			// Unexpected - it worked in our first AppDomain....?
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_EXCELDNALOADER, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return 0;
		}
	}
	else
	{
		hr = pSandbox->Load_3(bytesLoader, &pExcelDnaLoaderAssemblyInSandbox);
		if (FAILED(hr))
		{
			// Unexpected - it worked in our first AppDomain....?
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_EXCELDNALOADER, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return 0;
		}
	}


	CComPtr<_Type> pXlAddInType;
	hr = pExcelDnaLoaderAssemblyInSandbox->GetType_2(CComBSTR(L"ExcelDna.Loader.XlAddIn"), &pXlAddInType);
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

	if (!pAppDomain.IsEqualObject(pSandbox))
	{
		// Unload the loader AppDomain.
		pHost->UnloadDomain(pAppDomain);
	}

	// Keep references needed for later unload.
	pAppDomain_ForUnload = (IUnknown*)pSandbox.Detach();
	pHost_ForUnload = pHost.Detach();
	return initRetVal.boolVal == 0 ? false : true;
}

// XlLibraryUnload is only called if we are unloading the add-in via the add-in manager.
// Unload the AppDomain.
void XlLibraryUnload()
{
	if (pHost_ForUnload != NULL)
	{
		if (pAppDomain_ForUnload != NULL)
		{
			HRESULT hr = pHost_ForUnload->UnloadDomain(pAppDomain_ForUnload);
			pAppDomain_ForUnload->Release();
			pAppDomain_ForUnload = NULL;
			if (FAILED(hr))
			{
#if _DEBUG
				DebugBreak();
#endif
			}
		}
		//else
		//{
			// Unload according to the AppDomainId.
		//}

		pHost_ForUnload->Release();
		pHost_ForUnload = NULL;
	}
	// Also delete the temp .config file, if we made one.
	if (tempConfigFileName != "")
	{
		DeleteTempFile(tempConfigFileName);
		tempConfigFileName = "";
	}
}

// Try to get the right version of the CLR running.
HRESULT LoadClr(CString clrVersion, ICorRuntimeHost **ppHost)
{
	// Check whether the .Net 4+ MetaHost interfaces are present.
	// The checks here are according to this blog post: 
	// http://bradwilson.typepad.com/blog/2010/04/selecting-clr-version-from-unmanaged-host.html
	/*
	1. LoadLibrary mscoree
	2. GetProcAddress for CLRCreateInstance. If you get NULL, fall back to legacy path (CorBindToRuntimeEx)
	3. Call CLRCreateInstance to get ICLRMetaHost. If you get E_NOTIMPL, fall back to legacy path (same as above)
	4. Otherwise, party on the ICLRMetaHost you just got
	*/
	// If present, load the desired version using the new interfaces.
	// If not, check if we want .Net 4+, if so fail, else load old-style.

	HRESULT hr = E_FAIL;
	HMODULE hMscoree = NULL;
	CComPtr<ICLRMetaHost> pMetaHost;

	bool needNet40 = (clrVersion.CompareNoCase(L"v4.0") >= 0);
	bool needMetaHost = needNet40;

	hMscoree = LoadLibrary(L"mscoree.dll");
	if (hMscoree == 0)
	{
		// No .Net installed
		if (needNet40)
		{
				ShowMessage(IDS_MSG_HEADER_NEEDCLR40, 
					IDS_MSG_BODY_LOADMSCOREE, 
					IDS_MSG_FOOTER_ENSURECLR40 );
		}
		else
		{
				ShowMessage(IDS_MSG_HEADER_NEEDCLR20, 
					IDS_MSG_BODY_LOADMSCOREE, 
					IDS_MSG_FOOTER_ENSURECLR20 );
		}
		hr = E_FAIL;
	}
	else
	{
		pfnCLRCreateInstance CLRCreateInstance = (pfnCLRCreateInstance)GetProcAddress(hMscoree, "CLRCreateInstance");
		if (CLRCreateInstance == 0)
		{
			// Certainly no .Net 4 installed
			if (needMetaHost)
			{
				// We need .Net 4.0 but it is not installed
				ShowMessage(IDS_MSG_HEADER_NEEDCLR40, 
							IDS_MSG_BODY_NOCLRCREATEINSTANCE, 
							IDS_MSG_FOOTER_ENSURECLR40 );
				hr = E_FAIL;
			}
			else
			{
				// We need only .Net 2.0 runtime and cannot MetaHost.
				// Load .Net 2.0 with old code path
				hr = LoadClr20(ppHost);
			}
		}
		else
		{
			hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, (LPVOID*)&pMetaHost);
			if (FAILED(hr))
			{
				// MetaHost is not available, even though we have a new version of MSCorEE.dll
				// Certainly no .Net 4 installed
				if (needMetaHost)
				{
					// We need .Net 4.0 but it is not installed
					ShowMessage(IDS_MSG_HEADER_NEEDCLR40, 
								IDS_MSG_BODY_CLRCREATEINSTANCEFAILED, 
								IDS_MSG_FOOTER_ENSURECLR40 );
					hr = E_FAIL;
				}
				else
				{
					// We need only .Net 2.0 runtime and cannot MetaHost.
					// Load .Net 2.0 with old code path
					hr = LoadClr20(ppHost);
				}
			}
			else
			{
				// Yay! We have a metahost
				hr = LoadClrMeta(clrVersion, pMetaHost, ppHost);
			}
		}
		FreeLibrary(hMscoree);
	}
	return hr;
}

// Load the desired Clr version using .Net 4+ the MetaHost interfaces.
HRESULT LoadClrMeta(CString clrVersion, ICLRMetaHost* pMetaHost, ICorRuntimeHost **ppHost)
{
	// Even if we want to load .Net 2.0, we might need to multi-host since .Net 4.0 runtime
	// might also be loaded.

	HRESULT hr = E_FAIL;
	CComPtr<ICLRRuntimeInfo> pRuntimeInfo;
	bool needNet40 = (clrVersion.CompareNoCase(L"v4.0") >= 0);

	hr = pMetaHost->GetRuntime(clrVersion, IID_ICLRRuntimeInfo, (LPVOID*)&pRuntimeInfo);
	if (FAILED(hr))
	{
		// The version we ask for is not installed.
		// I.e. we want 2.0 but only 4.0 is installed.
		ShowMessage(IDS_MSG_HEADER_VERSIONLOADFAILED, 
					IDS_MSG_BODY_METAHOSTGETRUNTIMEFAILED, 
					IDS_MSG_FOOTER_ENSUREVERSION );
		hr = E_FAIL;
	}
	else
	{
		hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (LPVOID*)ppHost); 
		if (FAILED(hr))
		{
			// Not sure why this would happen???
			ShowMessage( needNet40 ? IDS_MSG_HEADER_NEEDCLR40 : IDS_MSG_HEADER_NEEDCLR20, 
						IDS_MSG_BODY_RUNTIMEGETINTERFACEFAILED, 
						IDS_MSG_FOOTER_UNEXPECTED );

			hr = E_FAIL;
		}
		else
		{
			hr = S_OK;
		}
	}
	return hr;
}

// Try to get the CLR 2.0 running - .Net 4+ MetaHost stuff not present.
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

// Create a new temp file with the given content.
// Most of this copied from CAtlTemporaryFile....
HRESULT CreateTempFile(void* pBuffer, DWORD nBufSize, CString& fileName)
{
		TCHAR szPath[_MAX_PATH]; 
		TCHAR tmpFileName[_MAX_PATH]; 
		DWORD dwRet = GetTempPath(_MAX_DIR, szPath);
		if (dwRet == 0)
		{
			// Couldn't find temporary path;
			return AtlHresultFromLastError();
		}
		else if (dwRet > _MAX_DIR)
		{
			return DISP_E_BUFFERTOOSMALL;
		}

		if (!GetTempFileName(szPath, _T("DNA"), 0, tmpFileName))
		{
			// Couldn't create temporary filename;
			return AtlHresultFromLastError();
		}
		tmpFileName[_countof(tmpFileName)-1]='\0';

		HANDLE hFile = ::CreateFile(
			tmpFileName,
			GENERIC_WRITE,
			0,		// No sharing - we'll write and close
			NULL,	// default security
			CREATE_ALWAYS,
			FILE_ATTRIBUTE_NOT_CONTENT_INDEXED | FILE_ATTRIBUTE_TEMPORARY,
			NULL);	// no template

		if (hFile == INVALID_HANDLE_VALUE)
			return AtlHresultFromLastError();

		DWORD nBytesWritten;
		BOOL writeOK = ::WriteFile(hFile, pBuffer, nBufSize, &nBytesWritten, NULL);
		if (!writeOK)
			return AtlHresultFromLastError();

		BOOL closeOK = ::CloseHandle(hFile);
		if (!closeOK)
			return AtlHresultFromLastError();

		fileName = tmpFileName;
		return S_OK;
}

HRESULT DeleteTempFile(CString fileName)
{
		BOOL deleteOK = ::DeleteFile(tempConfigFileName);
		if (!deleteOK)
			return AtlHresultFromLastError();
		
		return S_OK;
}

// LoaderInitialize is called when the .dll gets PROCESS_ATTACH
// First initialization comes here.
// For now we only store our own module handle.
void LoaderInitialize(HMODULE hModule)
{
	hModuleCurrent = hModule;
}

// LoaderUnload is called when the .dll gets PROCESS_DETACH.
// Last chance to clean up anything.
// We just delete the temp .config file if we created one.
void LoaderUnload()
{
	if (tempConfigFileName != "")
	{
		DeleteTempFile(tempConfigFileName);
		tempConfigFileName = "";
	}
}


// Decide what version of the CLR to load.
// returns E_FAIL if no dna file information, 
// else S_OK and clrVersion has a version string.

// Version is updated:
//	"v2.0" -> "v2.0.50727"
//	"v4.0" -> "v4.0.30319"

HRESULT ReadClrOptions(CString& clrVersion, bool& shadowCopyFiles)
{
	HRESULT hr;
	CString header;

	hr = ReadDnaHeader(header);	// Errors will be shown in there.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, clrVersion, shadowCopyFiles); // No errors yet.
		if (FAILED(hr))
		{
			// XML Parse error
			ShowMessage(IDS_MSG_HEADER_DNAPROBLEM, 
			IDS_MSG_BODY_DNAPARSEFAILED, 
			IDS_MSG_FOOTER_ENSUREDNAFILE,
			hr);

			return E_FAIL;
		}
		
		// Default version expansions
		if (clrVersion == L"v2.0") clrVersion = L"v2.0.50727";
		if (clrVersion == L"v4.0") clrVersion = L"v4.0.30319";

	}
	return hr;
}

HRESULT ParseDnaHeader(CString header, CString& runtimeVersion, bool& shadowCopyFiles)
{
	HRESULT hr;

	int rootTagStart = header.Find(L"<DnaLibrary");
	if (rootTagStart == -1)
	{
		// Parse error
		return E_FAIL;
	}

	int rootTagEnd = header.Find(L">", rootTagStart);
	if (rootTagEnd == -1)
	{
		// Parse error
		return E_FAIL;
	}

	CString rootTag = header.Mid(rootTagStart, rootTagEnd - rootTagStart + 1);

	// CONSIDER: Some checks, e.g. "v.X..."
	hr = ReadAttributeValue(rootTag, L"RuntimeVersion", runtimeVersion);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		runtimeVersion = CLR_VERSION_20;
		hr = S_OK;
	}

	CString shadowCopyFilesValue;
	hr = ReadAttributeValue(rootTag, L"ShadowCopyFiles", shadowCopyFilesValue);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		shadowCopyFiles = false;
		hr = S_OK;
	}
	else // attribute read OK
	{
		if (shadowCopyFilesValue.CompareNoCase(L"true") == 0)
			shadowCopyFiles = true;
		else
			shadowCopyFiles = false;
	}

	return hr;
}

// Returns	S_OK if the attribute was found and read into the attributeValue string.
//			S_FALSE if the attribute was not found at all
//			E_FAIL if there was an XML syntax error in the tag.
HRESULT ReadAttributeValue(CString tag, CString attributeName, CString& attributeValue)
{
	attributeName.Append(L"=");
	int attributeNameLength = attributeName.GetLength();

	int attributeNameStart = tag.Find(attributeName);
	if (attributeNameStart == -1)
	{
		return S_FALSE;
	}

	TCHAR quoteChar = tag[attributeNameStart + attributeNameLength];
	if (quoteChar != L'\'' && quoteChar != L'\"')
	{
		// XML syntax error - not a valid attribute.
		return E_FAIL;
	}

	int attributeValueStart = attributeNameStart + attributeNameLength + 1;
	int attributeValueEnd = tag.Find(quoteChar, attributeValueStart);
	if (attributeValueEnd == -1)
	{
		// XML syntax error - not a valid attribute.
		return E_FAIL;
	}
	attributeValue = tag.Mid(attributeValueStart, attributeValueEnd - attributeValueStart);
	return S_OK;
}

HRESULT ReadDnaHeader(CString& header)
{
	// We find the .dna file and load a 1k string from the file.
	// To locate the file:
	// 1. First check for packed __MAIN__ DNA resource,
	// 2. Else load file next to .xll file, 
	// Else E_FAIL.
	// This sequence matches the load sequence in ExcelDna.Integration.DnaLibrary.Initialize().
	// NOTE: __MAIN__ DNA resource can not currently be compressed.
	
	HRESULT hr = E_FAIL;
	const DWORD MAX_HEADER_LENGTH = 1024;
	DWORD headerLength = 0;
	BYTE headerBuffer[MAX_HEADER_LENGTH] ;

	HRSRC hResDna = FindResource(hModuleCurrent, L"__MAIN__", L"DNA");
	if (hResDna != NULL)
	{
		HGLOBAL hDna = LoadResource(hModuleCurrent, hResDna);
		DWORD sizeDna = SizeofResource(hModuleCurrent, hResDna);
		void* pDna = LockResource(hDna);
		headerLength = min(sizeDna, MAX_HEADER_LENGTH);
		CopyMemory(headerBuffer, pDna, headerLength);
	}
	else
	{
		CAtlFile dnaFile;
		CPath dnaPath(AddInFullPath());
		dnaPath.RenameExtension(L".dna");
		if (!dnaPath.FileExists())
		{
			ShowMessage(IDS_MSG_HEADER_DNANOTFOUND, 
			IDS_MSG_BODY_DNAPATHNOTEXIST, 
			IDS_MSG_FOOTER_ENSUREDNAFILE,
			hr);

			return E_FAIL;
		}
		hr = dnaFile.Create(dnaPath, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING);
		if (FAILED(hr))
		{
			ShowMessage(IDS_MSG_HEADER_DNAPROBLEM, 
			IDS_MSG_BODY_DNAOPENFAILED, 
			IDS_MSG_FOOTER_UNEXPECTED,
			hr);

			return E_FAIL;
		}
		hr = dnaFile.Read((LPVOID)headerBuffer, MAX_HEADER_LENGTH, headerLength);
		if (FAILED(hr))
		{

			ShowMessage(IDS_MSG_HEADER_DNAPROBLEM, 
			IDS_MSG_BODY_DNAOPENFAILED, 
			IDS_MSG_FOOTER_UNEXPECTED,
			hr);

			return E_FAIL;
		}
	}
	if (IsBufferUTF8(headerBuffer, headerLength))
	{
		header = UTF8toUTF16(CStringA((char*)headerBuffer, headerLength));
	}
	else
	{
		header = CString((wchar_t*)headerBuffer, headerLength);
	}
	return S_OK;
}

BOOL IsBufferUTF8(BYTE* buffer, DWORD bufferLength)
{
	// Only UTF-8 and UTF-16 is supported (here)
	// The check here is naive - does not read the xml processing instruction.
	// CONSIDER: Use WIN32 API function IsTextUnicode ?

	// Check for byte order marks.
	if (bufferLength < 3)
	{
		// Doesn't matter - will fail later.
		return true;
	}
	if (buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
	{
		// Standard UTF-8 BOM
		return true;
	}
	//if (buffer[0] == 0xFF && buffer[1] == 0xFE && buffer[2] == 0x00 && buffer[3] == 0x00)
	//{
	//	// UTF-32 LE
	//	return false;
	//}
	//if (buffer[0] == 0x00 && buffer[1] == 0x00 && buffer[2] == 0xFE && buffer[3] == 0xFF)
	//{
	//	// UTF-32 BE
	//	return false;
	//}
	if (buffer[0] == 0xFF && buffer[1] == 0xFE)
	{
		// UTF-16 LE
		return false;
	}
	if (buffer[0] == 0xFE && buffer[1] == 0xFF)
	{
		// UTF-16 BE
		return false;
	}
	// Might be ANSI or some other code page. Treated as UTF-8 here.
	return true;
}

// Snippet from http://www.codeproject.com/KB/string/utfConvert.aspx
CStringW UTF8toUTF16(const CStringA& utf8)
{
   CStringW utf16;
   int len = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
   if (len>1)
   { 
      wchar_t *ptr = utf16.GetBuffer(len-1);
      if (ptr) MultiByteToWideChar(CP_UTF8, 0, utf8, -1, ptr, len);
      utf16.ReleaseBuffer();
   }
   return utf16;
}
