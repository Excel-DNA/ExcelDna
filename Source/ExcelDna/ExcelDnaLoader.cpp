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

// ExcelDnaLoader.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"

//#include "CLRVer.h"
//#include "mscoree.h"
//#include "mscorlib_extract.h"
//#include "DetectDotNet.h"

#include "ExcelDna.h"
#include "ExcelDnaLoader.h"
#include "ExcelClrLoader.h"
#include "resource.h"
#define CountOf(x) sizeof(x)/sizeof(*x)

HMODULE hModuleCurrent;

//
//typedef HRESULT ( __stdcall *FLockClrVersionCallback ) ();





//
//typedef HRESULT (STDAPICALLTYPE *pCBRX)(        
//									LPWSTR pwszVersion,   
//									LPWSTR pwszBuildFlavor, 
//									DWORD flags,            
//									REFCLSID rclsid,      
//									REFIID riid,    
//									LPVOID* ppv );
//
//
//typedef HRESULT (STDAPICALLTYPE *pLCV) (
//    FLockClrVersionCallback hostCallback,
//    FLockClrVersionCallback *pBeginHostSetup,
//    FLockClrVersionCallback *pEndHostSetup
//);
//

			// This is the function pointer defintion for the shim API GetCorVersion.
// It has existed in mscoree.dll since v1.0, and will display the version of the runtime that is currently
// loaded into the process. If a CLR is not loaded into the process, it will load the latest version.
//typedef HRESULT (STDAPICALLTYPE *pGetCV)(LPWSTR szBuffer, 
//                                         DWORD cchBuffer,
//                                         DWORD* dwLength);
//

typedef HRESULT (STDAPICALLTYPE *pGetCORVersion)(LPWSTR pBuffer, 
                                         DWORD cchBuffer,
                                         DWORD* dwLength);

typedef HRESULT (STDAPICALLTYPE *pGetVersionFromProcess)(
										 HANDLE hProcess,
										 LPWSTR pBuffer, 
                                         DWORD cchBuffer,
                                         DWORD* dwLength);

//static FLockClrVersionCallback BeginHostSetup;
//static FLockClrVersionCallback EndHostSetup;

	//HRESULT STDAPICALLTYPE MyLockClrVersionCallback()
	//{
	//	ICorRuntimeHost *pHost = NULL;
	//    HMODULE MscoreeHandle = NULL;
	//    MscoreeHandle = LoadLibraryA("mscoree.dll");
	//	pCBRX CorBindToRuntimeExFunc = (pCBRX)GetProcAddress(MscoreeHandle, "CorBindToRuntimeEx");

	//	HRESULT resBegin = BeginHostSetup();
	//	HRESULT res = CorBindToRuntimeExFunc(L"v2.0.50727", L"wks", STARTUP_LOADER_SAFEMODE, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (PVOID*) &pHost);
	//	//HRESULT resStart = pHost->Start();
	//	HRESULT resEnd = EndHostSetup();
	//	// Do stuff here
	//	return S_OK;
	//}


	//_Assembly* STDAPICALLTYPE MyResolveEventHandler(_Object *pSender, _ResolveEventArgs *pResolveEventArgs)
	//{
	//	return NULL;
	//};



	bool XlLibraryInitialize(XlAddInExportInfo* pExportInfo)
	{
		HRESULT hr;
		CComPtr<ICorRuntimeHost> pHost;

		//hr = ExcelClrLoadByConfig(&pHost);
		//hr = ExcelClrLoadByConfigFile(&pHost);
		//hr = ExcelClrLoadDebug(&pHost);
		hr = ExcelClrLoad(&pHost);

		// ExcelClrLoad returns S_FALSE if a v2+ CLR is already loaded, 
		// S_OK if the CLR was loaded successfully, E_FAIL if a CLR could not be loaded.
		// ExcelClrLoad shows diagnostic MessageBoxes if needed.
		if (FAILED(hr) || pHost == NULL)
		{
			// Perhaps remember that we are not loaded?
			return 0;
		}
#ifdef _DEBUG
		MessageBox(NULL, L"ClrLoaded and ready.", L"ExcelDna Loader Diagnostics", 0);
#endif

		_TSCHAR FileName[MAX_PATH+1];
		DWORD dwResult = GetModuleFileName(hModuleCurrent, FileName, MAX_PATH);
		if (dwResult == 0)
		{
			MessageBox(NULL, L"Module file name could not be determined.", L"ExcelDna Loader Diagnostics", 0);
		}

		CPathT<CString> xllDirectory = FileName;
		xllDirectory.RemoveFileSpec();

		CComPtr<IUnknown> pAppDomainSetupUnk;
		hr = pHost->CreateDomainSetup(&pAppDomainSetupUnk);
		if (FAILED(hr) || pAppDomainSetupUnk == NULL)
		{
			MessageBox(NULL, L"AppDomainSetup could not be created.", L"ExcelDna Loader Diagnostics", 0);
			return 0;
		}

		CComQIPtr<IAppDomainSetup> pAppDomainSetup = pAppDomainSetupUnk;
		//IUnknown *pAppDomainSetupUnk = NULL;
		//hr = pHost->CreateDomainSetup(&pAppDomainSetupUnk);
		//if (FAILED(hr))
		//{
		//	MessageBox(NULL, L"AppDomainSetup could not be created.", L"ExcelDna Loader Diagnostics", 0);
		//	return 0;
		//}

		//IAppDomainSetup *pAppDomainSetup = NULL;
		//hr = pAppDomainSetupUnk->QueryInterface(IID_IAppDomainSetup, (void**)&pAppDomainSetup);
		//if (FAILED(hr))
		//{
		//	MessageBox(NULL, L"AppDomainSetup interface could not be retrieved.", L"ExcelDna Loader Diagnostics", 0);
		//	return 0;
		//}

		// TODO: Right path etc.
		hr = pAppDomainSetup->put_ApplicationBase(CComBSTR(xllDirectory));
		if (FAILED(hr))
		{
			MessageBox(NULL, L"ApplicationBase could not be set.", L"ExcelDna Loader Diagnostics", 0);
			return 0;
		}
		
		// TODO: Fix file name (and check?)
		CComBSTR configFileName = FileName;
		configFileName.Append(L".config");
		pAppDomainSetup->put_ConfigurationFile(configFileName);

		CComBSTR appDomainName = L"ExcelDna: ";
		appDomainName.Append(FileName);
		pAppDomainSetup->put_ApplicationName(appDomainName);

		IUnknown *pAppDomainUnk = NULL;
		hr = pHost->CreateDomainEx(appDomainName, pAppDomainSetupUnk, 0, &pAppDomainUnk);
		if (FAILED(hr) || pAppDomainUnk == NULL)
		{
			MessageBox(NULL, L"AppDomain could not be created.", L"ExcelDna Loader Diagnostics", 0);
			return 0;
		}

		CComQIPtr<_AppDomain> pAppDomain(pAppDomainUnk);

		// Load plan for ExcelDna.Loader:
		// Try AppDomain.Load with the name ExcelDna.Loader.
		// TODO: If it does not work, we will try to load from a known resource.

		CComPtr<_Assembly> pExcelDnaLoaderAssembly;
		hr = pAppDomain->Load_2(CComBSTR(L"ExcelDna.Loader"), &pExcelDnaLoaderAssembly);
		if (FAILED(hr) || pExcelDnaLoaderAssembly == NULL)
		{
// 			MessageBox(NULL, L"ExcelDna.Loader assembly could not be loaded - attempting resource load.", L"ExcelDna Loader Diagnostics", 0);


//			HMODULE hModule = GetModuleHandle(L"ExcelDnaLoader.xll");

#ifdef _DEBUG
			MessageBox(NULL, L"Loading ExcelDna.Loader from resources in: ", L"ExcelDna Loader Diagnostics", 0);
			MessageBox(NULL, FileName, L"ExcelDna Loader Diagnostics", 0);
#endif				
			HRSRC hResInfoLoader = FindResource(hModuleCurrent, L"EXCELDNA_LOADER", L"ASSEMBLY");
			if (hResInfoLoader == NULL)
			{
				MessageBox(NULL, L"ExcelDna_Loader Assembly could not be found in resources.", L"ExcelDna Loader Diagnostics", 0);
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
				MessageBox(NULL, L"Loader assembly could not be loaded from resource.", L"ExcelDna Loader Diagnostics", 0);
				return 0;
			}

			CComBSTR pFullName;
			hr = pExcelDnaLoaderAssembly->get_FullName(&pFullName);

			if (FAILED(hr))
			{
				MessageBox(NULL, L"Name for Loader assembly could not be determined.", L"ExcelDna Loader Diagnostics", 0);
				return 0;
			}
		}
		
		CComPtr<_Type> pXlLibraryType;
		hr = pExcelDnaLoaderAssembly->GetType_2(CComBSTR(L"ExcelDna.Loader.XlLibrary"), &pXlLibraryType);
		if (FAILED(hr) || pXlLibraryType == NULL)
		{
			MessageBox(NULL, L"XlLibrary type could not be loaded.", L"ExcelDna Loader Diagnostics", 0);
			return 0;
		}

		CComSafeArray<VARIANT> initArgs;
		initArgs.Add(CComVariant((INT32)pExportInfo));
		initArgs.Add(CComVariant((INT32)hModuleCurrent));
		initArgs.Add(CComVariant(FileName));
		CComVariant initRetVal;
		CComVariant target;
		hr = pXlLibraryType->InvokeMember_3(CComBSTR("Initialize"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, target, initArgs, &initRetVal);
		if (FAILED(hr))
		{
			MessageBox(NULL, L"XlLibrary.Initialize call failed.", L"ExcelDna Loader Diagnostics", 0);
			return 0;
		}
		
		return initRetVal.boolVal == 0 ? false : true;

//		_Assembly *pMscorlibAssembly = NULL;
//		hr = pAppDomain->Load_2(CComBSTR( L"mscorlib"), &pMscorlibAssembly);
//
//		_Type* pMarshalType = NULL;
//		hr = pMscorlibAssembly->GetType_2(CComBSTR("System.Runtime.InteropServices.Marshal"), &pMarshalType);
//
//		_Type *pResolveEventHandlerType = NULL;
//		hr = pMscorlibAssembly->GetType_2(CComBSTR("System.ResolveEventHandler"), &pResolveEventHandlerType);
//
//		//_MethodInfo *pMethodInfo = NULL;
//		//BindingFlags flags = (BindingFlags)(BindingFlags_Static | BindingFlags_Public);
//		//hr = pMarshalType->GetMethod_2(CComBSTR(L"GetDelegateForFunctionPointer"), flags, &pMethodInfo);
//
//		//CComSafeArray<VARIANT> parameters(2);
//		//CComVariant vObj;
//		////VARIANT vFunc;
//		////CComVariant vFunc;//((long)MyResolveEventHandler);
//		//CComVariant vType((IUnknown*)pResolveEventHandlerType);
//		//parameters[0] = (INT)0;
//		//parameters[1] = vType;
//		//CComVariant retval;
//		//hr = pMethodInfo->Invoke_3(vObj, parameters, &retval);
//
//		// // retrieve DISPID
//		//DISPID dispidGetDelegate;
//		//GUID iidNull = IID_NULL;
//		//CComBSTR NameGetDelegate(L"GetDelegateForFunctionPointer");
//		//hr = pMarshalType->GetIDsOfNames( &iidNull, (long)&NameGetDelegate.m_str, 1, LOCALE_SYSTEM_DEFAULT, (long)&dispidGetDelegate );
//		//if (FAILED(hr))
//		//{
//		//	MessageBox(NULL, L"GetDelegateForFunctionPointer DISPID could not be retrieved.", L"ExcelDna Loader Diagnostics", 0);
//		//	return 0;
//		//}
//
//		//// invoke method with two params on IDispatch
//		//CComVariant vDelegate;
//		//DISPPARAMS dispGetDelegateParams;
//		//memset( &dispGetDelegateParams, 0, sizeof( dispGetDelegateParams ) );
//		//dispGetDelegateParams.cArgs = 2;
//
//		//VARIANTARG* pvBufferGetDelegate = new VARIANTARG[ dispGetDelegateParams.cArgs ];
//		//dispGetDelegateParams.rgvarg = pvBufferGetDelegate;
//		//VariantInit( dispGetDelegateParams.rgvarg );
//		//dispGetDelegateParams.rgvarg[0].vt = VT_PTR;
//		//dispGetDelegateParams.rgvarg[0].byref = (void*)MyResolveEventHandler;
//		//dispGetDelegateParams.rgvarg[1].vt = VT_UNKNOWN;
//		//dispGetDelegateParams.rgvarg[1].punkVal = (IUnknown*)pResolveEventHandlerType;
//
//		//hr = pMarshalType->Invoke( dispidGetDelegate, &iidNull, 
//		//						LOCALE_SYSTEM_DEFAULT,
//		//						DISPATCH_METHOD, 
//		//						(long)&dispGetDelegateParams, 
//		//						(long)&vDelegate, NULL, NULL );
//
//		//_ResolveEventHandler *pResolveEventHandler = NULL;
//		//pAppDomain->add_AssemblyResolve(pResolveEventHandler);
//
//		CAtlFile *pFile = new CAtlFile();
//		hr = pFile->Create(L"C:\\Work\\ExcelDna\\ExcelDnaLoader\\ExcelDna.Loader\\bin\\Debug\\ExcelDna.Loader.dll", 
//						FILE_READ_DATA, 
//						FILE_SHARE_READ,
//						OPEN_EXISTING);
//		ULONGLONG nLen = 0;
//		hr = pFile->GetSize(nLen);
//		BYTE bytes[100000];
//		memset((void*)bytes, 0, CountOf(bytes));
//		DWORD nBytesRead = 0;
//		hr = pFile->Read(bytes, CountOf(bytes), nBytesRead);
//
//		CComSafeArray<BYTE> byteArray;
//		byteArray.Add(nBytesRead, bytes);
//
//		_Assembly *pAssembly = NULL;
////		hr = pAppDomain->Load_2(CComBSTR( L"ExcelDna.Loader, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" ), &pAssembly);
//		hr = pAppDomain->Load_3(byteArray, &pAssembly);
//		if (FAILED(hr))
//		{
//			MessageBox(NULL, L"Loader assembly could not be loaded.", L"ExcelDna Loader Diagnostics", 0);
//			return 0;
//		}
//
//		CComBSTR pFullName;
//		hr = pAssembly->get_FullName(&pFullName);
//
//		CComVariant vUnwrapped;
//		hr = pAssembly->CreateInstance(
//			CComBSTR( L"ExcelDna.Loader.EntryPoint" ),
//			&vUnwrapped);
//
//		//_ObjectHandle *pEntryPointHandle;
//		////hr = pAppDomain->CreateInstanceFrom(
//		////	CComBSTR( L"ExcelDna.Loader.dll"), 
//		////	CComBSTR( L"ExcelDna.Loader.EntryPoint" ), 
//		////	&pEntryPointHandle);
//		//hr = pAppDomain->CreateInstance(
//		//	pFullName, 
//		//	//CComBSTR( L"ExcelDna.Loader, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" ),
//		//	CComBSTR( L"ExcelDna.Loader.EntryPoint" ), 
//		//	&pEntryPointHandle);
//		//if (FAILED(hr))
//		//{
//		//	MessageBox(NULL, L"EntryPoint object could not be created.", L"ExcelDna Loader Diagnostics", 0);
//		//	return 0;
//		//}
//
//		//hr = pEntryPointHandle->Unwrap( &vUnwrapped );
//		if (FAILED(hr))
//		{
//			MessageBox(NULL, L"EntryPoint object could not be unwrapped.", L"ExcelDna Loader Diagnostics", 0);
//			return 0;
//		}
//		if ( vUnwrapped.vt != VT_UNKNOWN )
//		{
//			MessageBox(NULL, L"EntryPoint object returned unexpected interface.", L"ExcelDna Loader Diagnostics", 0);
//			return 0;
//		}
//
//		_Object *pEntryPointObject = NULL;
//		hr = vUnwrapped.punkVal->QueryInterface(IID__Object, (void**)&pEntryPointObject);
//		if (FAILED(hr))
//		{
//			MessageBox(NULL, L"EntryPoint dispatch interface not found.", L"ExcelDna Loader Diagnostics", 0);
//			return 0;
//		}
//
//		_Type* pEntryPointType = NULL;
//		hr = pAssembly->GetType_2(CComBSTR("ExcelDna.Loader.EntryPoint"), &pEntryPointType);
//		_MethodInfo *pMFPMethodInfo = NULL;
//		hr = pEntryPointType->GetMethod_2(CComBSTR(L"MakeFunctionPointer"), 
//			(BindingFlags)(BindingFlags_Public | BindingFlags_Instance),
//			&pMFPMethodInfo);
//
//		
//
//		CComVariant object((IUnknown*)pEntryPointObject);
//		CComVariant returnIntPtr;
//		CComSafeArray<VARIANT> parameters;
//		CComVariant ip;
//		ip.vt = VT_INT;
//		ip.intVal = 0;
//		parameters.Add(ip);
//		
//		hr = pMFPMethodInfo->Invoke_3(object, parameters, &returnIntPtr);
//		
//
//		 // retrieve DISPID
//		DISPID dispid;
//		CComBSTR Name(L"MakeFunctionPointer");
//		hr = pEntryPointObject->GetIDsOfNames( IID_NULL, &Name.m_str, 1, LOCALE_SYSTEM_DEFAULT, &dispid );
//		if (FAILED(hr))
//		{
//			MessageBox(NULL, L"MakeFunctionPointer DISPID could not be retrieved.", L"ExcelDna Loader Diagnostics", 0);
//			return 0;
//		}
//
//		//IntPtr IP;
//		//IP.m_value = (void*)MyResolveEventHandler;
//
//		// invoke method with one BSTR param on IDispatch
//		CComVariant vIntPtrResult;
//		DISPPARAMS dispEntryPointParams;
//		memset( &dispEntryPointParams, 0, sizeof( dispEntryPointParams ) );
//		dispEntryPointParams.cArgs = 1;
//
//		VARIANTARG* pvBuffer = new VARIANTARG[ dispEntryPointParams.cArgs ];
//		dispEntryPointParams.rgvarg = pvBuffer;
//		VariantInit( dispEntryPointParams.rgvarg );
//		dispEntryPointParams.rgvarg[0].vt = VT_INT;
//		dispEntryPointParams.rgvarg[0].intVal = (INT)345;
//		//dispEntryPointParams.rgvarg[1].vt = VT_UINT;
//		//dispEntryPointParams.rgvarg[1].uintVal = (UINT)&IP;
//		//dispEntryPointParams.rgvarg[0].vt = VT_UNKNOWN;
//		//dispEntryPointParams.rgvarg[0].punkVal = pResolveEventHandlerType;
//
//		hr = pEntryPointObject->Invoke( dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
//								DISPATCH_METHOD, 
//								&dispEntryPointParams, 
//								&vIntPtrResult, NULL, NULL );
//
//		//		 // retrieve DISPID
//		//DISPID dispid;
//		//CComBSTR Name(L"MakeDelegate");
//		//hr = pEntryPointObject->GetIDsOfNames( IID_NULL, &Name.m_str, 1, LOCALE_SYSTEM_DEFAULT, &dispid );
//		//if (FAILED(hr))
//		//{
//		//	MessageBox(NULL, L"LoadAddIn DISPID could not be retrieved.", L"ExcelDna Loader Diagnostics", 0);
//		//	return 0;
//		//}
//
//		//IntPtr IP;
//		//IP.m_value = (void*)MyResolveEventHandler;
//
//		//// invoke method with one BSTR param on IDispatch
//		//CComVariant vIgnoredResult;
//		//DISPPARAMS dispEntryPointParams;
//		//memset( &dispEntryPointParams, 0, sizeof( dispEntryPointParams ) );
//		//dispEntryPointParams.cArgs = 2;
//
//		//VARIANTARG* pvBuffer = new VARIANTARG[ dispEntryPointParams.cArgs ];
//		//dispEntryPointParams.rgvarg = pvBuffer;
//		//VariantInit( dispEntryPointParams.rgvarg );
//		////dispEntryPointParams.rgvarg[0].vt = VT_INT;
//		////dispEntryPointParams.rgvarg[0].intVal = (INT)345;
//		//dispEntryPointParams.rgvarg[1].vt = VT_UINT;
//		//dispEntryPointParams.rgvarg[1].uintVal = (UINT)&IP;
//		//dispEntryPointParams.rgvarg[0].vt = VT_UNKNOWN;
//		//dispEntryPointParams.rgvarg[0].punkVal = pResolveEventHandlerType;
//
//		//hr = pEntryPointObject->Invoke( dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
//		//						DISPATCH_METHOD, 
//		//						&dispEntryPointParams, 
//		//						&vIgnoredResult, NULL, NULL );
//
//		//		 // retrieve DISPID
//		//DISPID dispid;
//		//CComBSTR Name(L"LoadAddIn");
//		//hr = pEntryPointObject->GetIDsOfNames( IID_NULL, &Name.m_str, 1, LOCALE_SYSTEM_DEFAULT, &dispid );
//		//if (FAILED(hr))
//		//{
//		//	MessageBox(NULL, L"LoadAddIn DISPID could not be retrieved.", L"ExcelDna Loader Diagnostics", 0);
//		//	return 0;
//		//}
//
//
//		//// invoke method with one BSTR param on IDispatch
//		//CComVariant vIgnoredResult;
//		//DISPPARAMS dispEntryPointParams;
//		//memset( &dispEntryPointParams, 0, sizeof( dispEntryPointParams ) );
//		//dispEntryPointParams.cArgs = 1;
//
//		//VARIANTARG* pvBuffer = new VARIANTARG[ dispEntryPointParams.cArgs ];
//		//dispEntryPointParams.rgvarg = pvBuffer;
//		//VariantInit( dispEntryPointParams.rgvarg );
//		//dispEntryPointParams.rgvarg[0].vt = VT_I4;
//		//dispEntryPointParams.rgvarg[0].lVal = 12345;
//
//		//hr = pEntryPointObject->Invoke( dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
//		//						DISPATCH_METHOD, 
//		//						&dispEntryPointParams, 
//		//						&vIgnoredResult, NULL, NULL );
//
////		delete pvBuffer;
//		if (FAILED(hr))
//		{
//			MessageBox(NULL, L"LoadAddIn dispatch call failed.", L"ExcelDna Loader Diagnostics", 0);
//			return 0;
//		}
//
//
//
//		return 1;
//
//
//
//
//
//		// If .Net 2.0 or later is not installed on this machine....
//		//		Show message and exit.
//		DWORD processId;
//		HANDLE hProcess;
//		processId = GetCurrentProcessId();
//		hProcess =  OpenProcess(  PROCESS_QUERY_INFORMATION |
//                                    PROCESS_VM_READ,
//                                    FALSE, processId );
//
//	    HMODULE MscoreeHandle = NULL;
//	    MscoreeHandle = LoadLibraryA("mscoree.dll");
////		pGetCORVersion GetCORVersionFunc = (pGetCORVersion)GetProcAddress(MscoreeHandle, "GetCORVersion");
//		pGetVersionFromProcess GetVersionFromProcessFunc = (pGetVersionFromProcess)GetProcAddress(MscoreeHandle, "GetVersionFromProcess");
//
//		HRESULT hr;
//		WCHAR   VersionBuffer[MAX_PATH];
//		DWORD  cchVersionBuffer = MAX_PATH;
//		DWORD  dwVersionLength;
//
////		hr = GetCORVersionFunc(VersionBuffer, cchVersionBuffer, &dwVersionLength); // Get Length
//		hr = GetVersionFromProcessFunc(hProcess, VersionBuffer, cchVersionBuffer, &dwVersionLength); // Get Length
//		if (FAILED(hr))
//		{
//			*((int*)0) = 1;
//		}
//
//		CloseHandle(hProcess);
//		hProcess = 0;
//
//		return 1;		

		//




		//	//CDetectDotNet detect;

		//	//vector<string> CLRVersions;
		//	//detect.IsDotNetPresent();

		//	//detect.EnumerateCLRVersions(CLRVersions);	

		//
		//ClrVerInfo clrVerInfo;
		//InitializeClrVerInfo(clrVerInfo);
		//GetClrVerInfo(clrVerInfo);
		//if (!clrVerInfo.ClrInstalled)
		//{
		//	// No runtime is installed
		//	MessageBox(NULL, "The .Net framework is not installed on this system.\r\nThis Excel add-in requires .Net 2.0 or later.\r\nPlease download and install the .Net framework from Microsoft.\r\nThe add-in will not continue loading.", "ExcelDna Loader", 0);
		//	return 0;
		//}

		////if (!clrVerInfo.v2PlusInstalled)
		////{
		////	// No runtime is installed
		////	MessageBox(NULL, L"Version 2.0 of the .Net framework is not installed on this system.\r\nThis Excel add-in requires .Net 2.0 or later.\r\nPlease download and install the .Net framework from Microsoft.\r\nThe add-in will not continue loading.", L"ExcelDna Loader", 0);
		////	return 0;
		////}

		//// Attempt to load the latest version of the CLR.
		////MessageBox(NULL, L"Checking if the .Net runtime is already loaded...", L"ExcelDna Loader", 0);


	 //   HMODULE MscoreeHandle = NULL;
	 //   MscoreeHandle = LoadLibraryA("mscoree.dll");
		//pCBRX CorBindToRuntimeExFunc = (pCBRX)GetProcAddress(MscoreeHandle, "CorBindToRuntimeEx");
		//pLCV LockClrVersionFunc = (pLCV)GetProcAddress(MscoreeHandle, "LockClrVersion");

		//HRESULT lcvResult = LockClrVersionFunc(MyLockClrVersionCallback, &BeginHostSetup, &EndHostSetup);

		//ICorRuntimeHost *pHost = NULL;
		////HRESULT res = CorBindToRuntimeExFunc(L"v8.0.50727", L"wks", STARTUP_LOADER_SAFEMODE, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (PVOID*) &pHost);
		//HRESULT res = CorBindToRuntimeExFunc(NULL, NULL, NULL, CLSID_CorRuntimeHost, IID_ICorRuntimeHost, (PVOID*) &pHost);

		//if (SUCCEEDED(res))
		//{
		//	pGetCV GetCorVersionFunc = (pGetCV)GetProcAddress(MscoreeHandle, "GetCORVersion");
		//	WCHAR Version[50];
		//	DWORD VersionNumChars = NumItems(Version);

		//	HRESULT hr = GetCorVersionFunc(Version, VersionNumChars, &VersionNumChars);


		//	res = pHost->Start();
		//}

		//return 0;
	}

void SetCurrentModule(HMODULE hModule)
{
	hModuleCurrent = hModule;
}
