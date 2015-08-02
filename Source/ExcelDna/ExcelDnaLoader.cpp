//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.


// TODO: Change to the .Net 2.0+ hosting interface IClrRuntimeHost, 
// so that we can set the safe AppDomain flags when loading.

#include "stdafx.h"
#include <string>
#include <tchar.h>
#include "resource.h"
#include "MiscUtils.h"
#include "DetectFx.h"
#include "ExcelDna.h"
#include "ExcelDnaLoader.h"

#define CountOf(x) sizeof(x)/sizeof(*x)
const int MAX_MSG = 1024;
const std::wstring CLR_VERSION_20 = L"v2.0.50727";

static HMODULE hModuleCurrent;
// These don't use ATL classes to give us explicit control over when CLR is called
static IUnknown* pAppDomain_ForUnload = NULL;
static ICorRuntimeHost* pHost_ForUnload = NULL;
// Temp file to be used if we need to write .config from resources.
static TempFileHolder tempConfig;

// Forward declarations for functions defined in this file.
HRESULT LoadClr(std::wstring clrVersion, ICorRuntimeHost **ppHost);
HRESULT LoadClrMeta(std::wstring clrVersion, ICLRMetaHost* pMetaHost, ICorRuntimeHost **ppHost);
HRESULT LoadClr20(ICorRuntimeHost **ppHost);

_COM_SMARTPTR_TYPEDEF(ICorRuntimeHost, IID_ICorRuntimeHost);
EXTERN_GUID(IID__Assembly,	0x17156360, 0x2f1a, 0x384a, 0xbc, 0x52, 0xfd, 0xe9, 0x3c, 0x21, 0x5c, 0x5b);
_COM_SMARTPTR_TYPEDEF(_Assembly, IID__Assembly);
_COM_SMARTPTR_TYPEDEF(_AppDomain, IID__AppDomain);
_COM_SMARTPTR_TYPEDEF(IAppDomainSetup, IID_IAppDomainSetup);
_COM_SMARTPTR_TYPEDEF(_Type, IID__Type);
_COM_SMARTPTR_TYPEDEF(ICLRMetaHost, IID_ICLRMetaHost);
_COM_SMARTPTR_TYPEDEF(ICLRRuntimeInfo, IID_ICLRRuntimeInfo);

HRESULT LoadAppDomain(ICorRuntimeHostPtr pHost, std::wstring addInFullPath, bool createSandboxedAppDomain, bool shadowCopyFiles, _AssemblyPtr& loaderAssembly, _AppDomainPtr& addInAppDomain, bool& unloadAppDomain);
HRESULT LoadLoaderIntoAppDomain(_AppDomainPtr& pAppDomain, _AssemblyPtr& pLoaderAssembly, bool forceFromBytes);
void ShowMessage(int headerId, int bodyId, int footerId, HRESULT hr = S_OK);
HRESULT CreateTempFile(void* pBuffer, DWORD nBufSize, std::wstring& fileName);
HRESULT DeleteTempFile(std::wstring fileName);

HRESULT GetClrOptions(std::wstring& clrVersion, bool& shadowCopyFiles, bool& createSandboxedAppDomain);
HRESULT GetDnaHeader(bool showErrors, std::wstring& header);
HRESULT ParseDnaHeader(std::wstring header, std::wstring& addInName, std::wstring& runtimeVersion, bool& shadowCopyFiles, std::wstring& createSandboxedAppDomain);
HRESULT GetAttributeValue(std::wstring tag, std::wstring attributeName, std::wstring& attributeValue);

BOOL IsRunningOnCluster();

BOOL IsBufferUTF8(BYTE* buffer, DWORD bufferLength);
std::wstring UTF8toUTF16(const std::string& utf8);

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
	ICorRuntimeHostPtr pHost;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	bool createSandboxedAppDomain;
	
	hr = GetClrOptions(clrVersion, shadowCopyFiles, createSandboxedAppDomain);
	if (FAILED(hr))
	{
		// SelectClrVersion shows diagnostic MessageBoxes if needed.
		// Perhaps remember that we are not loaded?
		return 0;
	}
#ifdef _M_X64
	bool allowedVersion = CompareNoCase(clrVersion, L"v4.0") >= 0;
	if (!allowedVersion)
	{
		ShowMessage(IDS_MSG_HEADER_64NET4, 
					IDS_MSG_BODY_64NET4,
					IDS_MSG_FOOTER_ENSUREVERSION,
					hr);
		return 0;
	}
#endif
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

	// Load (or find) the AppDomain that will contain the add-in
	std::wstring addInFullPath = GetAddInFullPath();
	_AppDomainPtr pAppDomain;
	_AssemblyPtr pLoaderAssembly;
	bool unloadAppDomain;

	hr = LoadAppDomain(pHost, addInFullPath, createSandboxedAppDomain, shadowCopyFiles, pLoaderAssembly, pAppDomain, unloadAppDomain);
	if (FAILED(hr))
	{
		// Message already shown by LoadAppDomain
		return 0;
	}

	_TypePtr pXlAddInType;
	hr = pLoaderAssembly->GetType_2(_bstr_t(L"ExcelDna.Loader.XlAddIn"), &pXlAddInType);
	if (FAILED(hr) || pXlAddInType == NULL)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_XLADDIN, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return 0;
	}

	SafeVariantArray initArgs(3);
	initArgs.lock();
#ifndef _M_X64
	initArgs.setElement(0, (INT32)pExportInfo);
	initArgs.setElement(1, (INT32)hModuleCurrent);
#else
	initArgs.setElement(0, (INT64)pExportInfo);
	initArgs.setElement(1, (INT64)hModuleCurrent);
#endif
	initArgs.setElement(2, SysAllocStringLen(addInFullPath.c_str(), static_cast<UINT>(addInFullPath.length())));
	initArgs.unlock();
	_variant_t initRetVal;
	_variant_t target;
#ifndef _M_X64
	hr = pXlAddInType->InvokeMember_3(_bstr_t("Initialize32"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, target, initArgs, &initRetVal);
#else
	hr = pXlAddInType->InvokeMember_3(_bstr_t("Initialize64"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, target, initArgs, &initRetVal);
#endif
	if (FAILED(hr) || initRetVal.boolVal == 0)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_XLADDININIT, 
					IDS_MSG_FOOTER_CHECKINTEGRATION,
					hr);
		return 0;
	}

	// Keep references needed for later host reference unload.
	if (unloadAppDomain)
	{
		pAppDomain_ForUnload = (IUnknown*)pAppDomain.Detach();
	}
	pHost_ForUnload = pHost.Detach();

	return initRetVal.boolVal == 0 ? false : true;
}

// XlLibraryUnload is only called if we are unloading the add-in via the add-in manager,
// or when the add-in is re-loaded.
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
	tempConfig.destroy();
}

// Try to get the right version of the CLR running.
HRESULT LoadClr(std::wstring clrVersion, ICorRuntimeHost **ppHost)
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
	ICLRMetaHostPtr pMetaHost;

	bool needNet40 = (CompareNoCase(clrVersion, L"v4.0") >= 0);
	bool needMetaHost = needNet40;

	hMscoree = LoadLibrary(L"mscoree.dll");
	if (hMscoree == 0)
	{
		// No .Net installed
		// CONSIDER: Doing explicit checking according to http://support.microsoft.com/kb/318785
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
								IDS_MSG_FOOTER_ENSURECLR40, 
								hr );
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
HRESULT LoadClrMeta(std::wstring clrVersion, ICLRMetaHost* pMetaHost, ICorRuntimeHost **ppHost)
{
	// Even if we want to load .Net 2.0, we might need to multi-host since .Net 4.0 runtime
	// might also be loaded.

	HRESULT hr = E_FAIL;
	ICLRRuntimeInfoPtr pRuntimeInfo;
	bool needNet40 = (CompareNoCase(clrVersion, L"v4.0") >= 0);

	hr = pMetaHost->GetRuntime(clrVersion.c_str(), IID_ICLRRuntimeInfo, (LPVOID*)&pRuntimeInfo);
	if (FAILED(hr))
	{
		// The version we ask for is not installed.
		// I.e. we want 2.0 but only 4.0 is installed.
		ShowMessage(IDS_MSG_HEADER_VERSIONLOADFAILED, 
					IDS_MSG_BODY_METAHOSTGETRUNTIMEFAILED, 
					IDS_MSG_FOOTER_ENSUREVERSION, 
					hr);
		hr = E_FAIL;
	}
	else
	{
		// Test for mixed file binding issue - try before or after the GetInterface call?
		// Discussed here: http://exceldna.codeplex.com/discussions/253974
		//		HRESULT hrbind = pRuntimeInfo->BindAsLegacyV2Runtime();
		hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (LPVOID*)ppHost); 
		if (FAILED(hr))
		{
			// Not sure why this would happen???
			ShowMessage( needNet40 ? IDS_MSG_HEADER_NEEDCLR40 : IDS_MSG_HEADER_NEEDCLR20, 
						IDS_MSG_BODY_RUNTIMEGETINTERFACEFAILED, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);

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
								IDS_MSG_FOOTER_ENSURECLR20,
								hr);
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

HRESULT LoadAppDomain(ICorRuntimeHostPtr pHost, std::wstring addInFullPath, bool createSandboxedAppDomain, bool shadowCopyFiles, _AssemblyPtr& pLoaderAssembly , _AppDomainPtr& pAppDomain, bool& unloadAppDomain)
{
	HRESULT hr;
	std::wstring xllDirectory = addInFullPath;
	RemoveFileSpecFromPath(xllDirectory);
	unloadAppDomain = false;

	if (IsRunningOnCluster())
	{
		// Need to load into default AppDomain due to configuration issues of the cluster host.
		IUnknown *pAppDomainUnk = NULL;
		hr = pHost->CurrentDomain(&pAppDomainUnk);
		if (FAILED(hr) || pAppDomainUnk == NULL)
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_APPDOMAIN, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
		
			return E_FAIL;
		}
		// Assignment does QueryInterface
		pAppDomain = pAppDomainUnk;
		hr = LoadLoaderIntoAppDomain(pAppDomain, pLoaderAssembly, /*forceFromBytes=*/ true);
		if (FAILED(hr))
		{
			// Already showed error message there.
			return E_FAIL;
		}
		// Since we loaded into the default domain, don't unload the AppDomain later.
		unloadAppDomain = false;
		return S_OK;
	}
	// End of RunningOnCluster path.

	// Create and populate AppDomainSetup
	IUnknownPtr pAppDomainSetupUnk;
	hr = pHost->CreateDomainSetup(&pAppDomainSetupUnk);
	if (FAILED(hr) || pAppDomainSetupUnk == NULL)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_APPDOMAINSETUP, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return E_FAIL;
	}

	IAppDomainSetupPtr pAppDomainSetup = pAppDomainSetupUnk;
	hr = pAppDomainSetup->put_ApplicationBase(_bstr_t(xllDirectory.c_str()));
	if (FAILED(hr))
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_APPLICATIONBASE, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return E_FAIL;
	}

	hr = pAppDomainSetup->put_ShadowCopyFiles(_bstr_t(shadowCopyFiles ? L"true" : L"false"));
	if (FAILED(hr))
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_SHADOWCOPYFILES, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return E_FAIL;
	}

	// AppDomainSetup.ApplicationName = "Excel-DNA: c:\MyAddins\MyAddIn.xll";
	_bstr_t appDomainName((std::wstring(L"Excel-DNA: ") + addInFullPath).c_str());
	pAppDomainSetup->put_ApplicationName(appDomainName);


	// Check if a .config file exists next to the .xll as MyAddIn.xll.config. Use it if it exists.
	std::wstring configFileName = addInFullPath + L".config";
	if (FileExists(configFileName.c_str()))
	{
		pAppDomainSetup->put_ConfigurationFile(_bstr_t(configFileName.c_str()));
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

			std::wstring tempConfigFileName;
			hr = CreateTempFile(pConfig, sizeConfig, tempConfigFileName);
			if (SUCCEEDED(hr))
			{
				pAppDomainSetup->put_ConfigurationFile(_bstr_t(tempConfigFileName.c_str()));
			}
			// tempConfigFile will be deleted after the AppDomain has been unloaded.
		}
		else
		{
			// No config file - no problem.
		}
	}

	IUnknown *pAppDomainUnk = NULL;
	hr = pHost->CreateDomainEx(appDomainName, pAppDomainSetupUnk, 0, &pAppDomainUnk);
	if (FAILED(hr) || pAppDomainUnk == NULL)
	{
		ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
					IDS_MSG_BODY_APPDOMAIN, 
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
		return E_FAIL;
	}

	pAppDomain = pAppDomainUnk;

	hr = LoadLoaderIntoAppDomain(pAppDomain, pLoaderAssembly, /*forceLoadFromBytes=*/ false);
	if (FAILED(hr))
	{
		// Already showed message.
		return E_FAIL;
	}
	
	if (createSandboxedAppDomain)
	{
		_TypePtr pAppDomainHelperType;
		hr = pLoaderAssembly->GetType_2(_bstr_t(L"ExcelDna.Loader.AppDomainHelper"), &pAppDomainHelperType);
		if (FAILED(hr) || pAppDomainHelperType == NULL)
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_XLADDIN, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return E_FAIL;
		}

		_variant_t  sbRetVal;
		_variant_t  sbTarget;
		hr = pAppDomainHelperType->InvokeMember_3(_bstr_t("CreateFullTrustSandbox"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, sbTarget, NULL, &sbRetVal);
		if (FAILED(hr))
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_XLADDININIT, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return E_FAIL;
		}

		_AppDomainPtr pSandbox(sbRetVal.punkVal);
		if (!IsEqualObject(pAppDomain, pSandbox))
		{
			// Unload the loader AppDomain.
			pLoaderAssembly.Release();
			pHost->UnloadDomain(pAppDomain);
			pAppDomain.Release();
			pAppDomain = pSandbox;

			//  Sort out the LoaderAssembly in the sandbox.
			hr = LoadLoaderIntoAppDomain(pAppDomain, pLoaderAssembly, /*forceLoadFromBytes=*/ false);
			if (FAILED(hr))
			{
				// Already showed message.
				return E_FAIL;
			}
		}
	}
	unloadAppDomain = true;
	return S_OK;
}

HRESULT LoadLoaderIntoAppDomain(_AppDomainPtr& pAppDomain, _AssemblyPtr& pLoaderAssembly, bool forceFromBytes)
{
	HRESULT hr;
	
	if (!forceFromBytes)
	{
		// Try regular load first 
		hr = pAppDomain->Load_2(_bstr_t(L"ExcelDna.Loader"), &pLoaderAssembly);
	}

	if (forceFromBytes || FAILED(hr) || pLoaderAssembly == NULL)
	{
		// Now try from resource bytes
		HRSRC hResInfoLoader = FindResource(hModuleCurrent, L"EXCELDNA.LOADER", L"ASSEMBLY");
		if (hResInfoLoader == NULL)
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_MISSINGEXCELDNALOADER, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return E_FAIL;
		}
		HGLOBAL hLoader = LoadResource(hModuleCurrent, hResInfoLoader);
		void* pLoader = LockResource(hLoader);
		ULONG sizeLoader = (ULONG)SizeofResource(hModuleCurrent, hResInfoLoader);
		
		SafeByteArray loaderBytes(pLoader, sizeLoader);

		hr = pAppDomain->Load_3(loaderBytes, &pLoaderAssembly);
		if (FAILED(hr))
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_EXCELDNALOADER, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return E_FAIL;
		}

		if (FAILED(hr))
		{
			ShowMessage(IDS_MSG_HEADER_APPDOMAIN, 
						IDS_MSG_BODY_EXCELDNALOADERNAME, 
						IDS_MSG_FOOTER_UNEXPECTED,
						hr);
			return E_FAIL;
		}
	}
	return S_OK;
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
		wchar_t className[11];
		DWORD count = RealGetWindowClass(hwnd, className, 10);
		if (_tcsncmp(className, L"XLMAIN", 6))
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
	if (IsRunningOnCluster())
	{
		// TODO: Consider what to do in cluster context?
		return;
	}

	HWND hwndExcel = FindCurrentExcelWindow();
	try
	{
		std::wstring  addInFullPath = GetAddInFullPath();
		std::wstring  addInFileName = addInFullPath;
		StripPath(addInFileName);

		std::wstring msgTitle = FormatString(LoadStringFromResource(hModuleCurrent, IDS_MSG_TITLE), addInFileName.c_str());

		std::wstring header = LoadStringFromResource(hModuleCurrent, headerId);
		std::wstring body = LoadStringFromResource(hModuleCurrent, bodyId);
		std::wstring footer = LoadStringFromResource(hModuleCurrent, footerId);

		std::wstring hresult = L"";
		if (hr != S_OK)
		{
			_com_error error(hr);
			hresult = FormatString(LoadStringFromResource(hModuleCurrent, IDS_MSG_HRESULT), error.ErrorMessage());
		}

		std::wstring msg = FormatString(LoadStringFromResource(hModuleCurrent, IDS_MSG_TEMPLATE), header.c_str(), body.c_str(), footer.c_str(), hresult.c_str(), addInFullPath.c_str());
		MessageBox(hwndExcel, msg.c_str(), msgTitle.c_str(), MB_ICONEXCLAMATION);
	}
	catch (...)
	{
		ShowMessageError(hwndExcel);
	}
}

std::wstring GetAddInFullPath()
{
	wchar_t buffer[MAX_PATH];
	DWORD count = GetModuleFileName(hModuleCurrent, buffer, MAX_PATH);
	return std::wstring(buffer);
}

// CONSIDER: DELETE_FILE_ON_CLOSE should actually work?
//           http://www.drdobbs.com/article/print?articleId=184416443&siteSectionName=windows
// Create a new temp file with the given content.
// Most of this copied from CAtlTemporaryFile....
HRESULT CreateTempFile(void* pBuffer, DWORD nBufSize, std::wstring& fileName)
{
		TCHAR szPath[_MAX_PATH]; 
		TCHAR tmpFileName[_MAX_PATH]; 
		DWORD dwRet = GetTempPath(_MAX_DIR, szPath);
		if (dwRet == 0)
		{
			// Couldn't find temporary path;
			return HResultFromLastError();
		}
		else if (dwRet > _MAX_DIR)
		{
			return DISP_E_BUFFERTOOSMALL;
		}

		if (!GetTempFileName(szPath, _T("DNA"), 0, tmpFileName))
		{
			// Couldn't create temporary filename;
			return HResultFromLastError();
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
			return HResultFromLastError();

		DWORD nBytesWritten;
		BOOL writeOK = ::WriteFile(hFile, pBuffer, nBufSize, &nBytesWritten, NULL);
		if (!writeOK)
			return HResultFromLastError();

		BOOL closeOK = ::CloseHandle(hFile);
		if (!closeOK)
			return HResultFromLastError();

		fileName = tmpFileName;
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
	tempConfig.destroy();
}

BOOL IsRunningOnCluster()
{
	// Our check is to see if the current process is called Excel.exe.
	// Hopefully this doen't change soon.
	
	TCHAR hostPathName[MAX_PATH];
	DWORD count = GetModuleFileName(NULL, hostPathName, MAX_PATH);

	std::wstring hostPath = hostPathName;
	StripPath(hostPath);

	if (CompareNoCase(hostPath, L"EXCEL.EXE") == 0)
	{
		return false;
	}
	
	return true;
}

HRESULT GetAddInName(std::wstring& addInName)
{
	HRESULT hr;
	std::wstring header;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	std::wstring createSandboxedAppDomainValue;

	hr = GetDnaHeader(false, header);	// Don't show errors here.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue); // No errors yet.
		if (FAILED(hr))
		{
			return E_FAIL;
		}
		if (addInName.empty())
		{
			std::wstring xllPath(GetAddInFullPath());
			StripPath(xllPath);
			RemoveExtension(xllPath);
			addInName = xllPath;
		}
	}
	return hr;
}

// Decide what version of the CLR to load.
// returns E_FAIL if no dna file information, 
// else S_OK and clrVersion has a version string.

// Version is updated:
//	"v2.0" -> "v2.0.50727"
//	"v4.0" -> "v4.0.30319"

HRESULT GetClrOptions(std::wstring& clrVersion, bool& shadowCopyFiles, bool& createSandboxedAppDomain)
{
	HRESULT hr;
	std::wstring header;
	std::wstring addInName;
	std::wstring createSandboxedAppDomainValue;

	hr = GetDnaHeader(true, header);	// Errors will be shown in there.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue); // No errors yet.
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

		// Default sandboxedAppDomain options
		if (CompareNoCase(createSandboxedAppDomainValue, L"true") == 0)
		{
			createSandboxedAppDomain = true;
		}
		else if (CompareNoCase(createSandboxedAppDomainValue, L"false") == 0)
		{
			createSandboxedAppDomain = false;
		}
		else
		{
			// Default => true under .NET >= 4.0, else false
			if (CompareNoCase(clrVersion, L"v4.0") >= 0)
				createSandboxedAppDomain = true;
			else
				createSandboxedAppDomain = false;
		}
	}
	return hr;
}

HRESULT ParseDnaHeader(std::wstring header, std::wstring& addInName, std::wstring& runtimeVersion, bool& shadowCopyFiles, std::wstring& createSandboxedAppDomain)
{
	HRESULT hr;

	size_t rootTagStart = header.find(L"<DnaLibrary");
	if (rootTagStart == -1)
	{
		// Parse error
		return E_FAIL;
	}

	size_t rootTagEnd = header.find(L">", rootTagStart);
	if (rootTagEnd == -1)
	{
		// Parse error
		return E_FAIL;
	}

	std::wstring rootTag = header.substr(rootTagStart, rootTagEnd - rootTagStart + 1);

	// CONSIDER: Some checks, e.g. "v.X..."
	hr = GetAttributeValue(rootTag, L"RuntimeVersion", runtimeVersion);
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

	std::wstring shadowCopyFilesValue;
	hr = GetAttributeValue(rootTag, L"ShadowCopyFiles", shadowCopyFilesValue);
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
		if (CompareNoCase(shadowCopyFilesValue, L"true") == 0)
			shadowCopyFiles = true;
		else
			shadowCopyFiles = false;
	}

	hr = GetAttributeValue(rootTag, L"CreateSandboxedAppDomain", createSandboxedAppDomain);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		createSandboxedAppDomain = L"";
		hr = S_OK;
	}

	hr = GetAttributeValue(rootTag, L"Name", addInName);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		addInName = L"";
		hr = S_OK;
	}
	return hr;
}

// Returns	S_OK if the attribute was found and read into the attributeValue string.
//			S_FALSE if the attribute was not found at all
//			E_FAIL if there was an XML syntax error in the tag.
// TODO: There is a bug here - I don't check the character before attributeName starts, so I also match XXXName="NotMyName"
//		 For not the .dna schema does not define any conflicts here, but it's not great.
HRESULT GetAttributeValue(std::wstring tag, std::wstring attributeName, std::wstring& attributeValue)
{
	attributeName += L"=";
	size_t attributeNameLength = attributeName.size();

	size_t attributeNameStart = tag.find(attributeName);
	if (attributeNameStart == -1)
	{
		return S_FALSE;
	}

	wchar_t quoteChar = tag[attributeNameStart + attributeNameLength];
	if (quoteChar != L'\'' && quoteChar != L'\"')
	{
		// XML syntax error - not a valid attribute.
		return E_FAIL;
	}

	size_t attributeValueStart = attributeNameStart + attributeNameLength + 1;
	size_t attributeValueEnd = tag.find(quoteChar, attributeValueStart);
	if (attributeValueEnd == -1)
	{
		// XML syntax error - not a valid attribute.
		return E_FAIL;
	}
	attributeValue = tag.substr(attributeValueStart, attributeValueEnd - attributeValueStart);
	return S_OK;
}

HRESULT GetDnaHeader(bool showErrors, std::wstring& header)
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
		SafeFile dnaFile;
		std::wstring dnaPath(GetAddInFullPath());
		RenameExtension(dnaPath, L".dna");
		if (!FileExists(dnaPath.c_str()))
		{
			if (showErrors)
			{
				ShowMessage(IDS_MSG_HEADER_DNANOTFOUND, 
				IDS_MSG_BODY_DNAPATHNOTEXIST, 
				IDS_MSG_FOOTER_ENSUREDNAFILE,
				hr);
			}
			return E_FAIL;
		}
		hr = dnaFile.Create(dnaPath, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING);
		if (FAILED(hr))
		{
			if (showErrors)
			{
				ShowMessage(IDS_MSG_HEADER_DNAPROBLEM, 
				IDS_MSG_BODY_DNAOPENFAILED, 
				IDS_MSG_FOOTER_UNEXPECTED,
				hr);
			}
			return E_FAIL;
		}
		hr = dnaFile.Read((LPVOID)headerBuffer, MAX_HEADER_LENGTH, headerLength);
		if (FAILED(hr))
		{
			if (showErrors)
			{
				ShowMessage(IDS_MSG_HEADER_DNAPROBLEM, 
				IDS_MSG_BODY_DNAOPENFAILED, 
				IDS_MSG_FOOTER_UNEXPECTED,
				hr);
			}
			return E_FAIL;
		}
	}
	if (IsBufferUTF8(headerBuffer, headerLength))
	{
		header = UTF8toUTF16(std::string((char*)headerBuffer, headerLength));
	}
	else
	{
		header = std::wstring((wchar_t*)headerBuffer, headerLength);
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

//// Snippet from http://www.codeproject.com/KB/string/utfConvert.aspx
//CStringW UTF8toUTF16(const CStringA& utf8)
//{
//   CStringW utf16;
//   int len = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
//   if (len>1)
//   { 
//      wchar_t *ptr = utf16.GetBuffer(len-1);
//      if (ptr) MultiByteToWideChar(CP_UTF8, 0, utf8, -1, ptr, len);
//      utf16.ReleaseBuffer();
//   }
//   return utf16;
//}
