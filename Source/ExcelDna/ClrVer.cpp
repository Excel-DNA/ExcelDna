// This file adapted from the CLRVer sample in the .NET Framework SDK.

#include "stdafx.h"
#include <stdio.h>
#include <mscoree.h>
#include <windows.h>
#include <Tlhelp32.h>

#define _UNICODE

#include "CLRVer.h"

#define NumItems(x) sizeof(x)/sizeof(*x)

// This is the function pointer defintion for the shim API GetRequestedRuntimeInfoInfo.
// It has existed in mscoree.dll since v1.1, and in v2.0 it was modified to take "runtimeInfoFlags"
// which allow us to get even more information.
typedef HRESULT (STDAPICALLTYPE *PGetRRI)(LPCWSTR pExe, 
                                          LPCWSTR pwszVersion,
                                          LPCWSTR pConfigurationFile, 
                                          DWORD startupFlags,
                                          DWORD runtimeInfoFlags, 
                                          LPWSTR pDirectory, 
                                          DWORD dwDirectory, 
                                          DWORD *dwDirectoryLength, 
                                          LPWSTR pVersion, 
                                          DWORD cchBuffer, 
                                          DWORD* dwlength);


// This is the function pointer defintion for the shim API GetCorVersion.
// It has existed in mscoree.dll since v1.0, and will display the version of the runtime that is currently
// loaded into the process. If a CLR is not loaded into the process, it will load the latest version.
typedef HRESULT (STDAPICALLTYPE *PGetCV)(LPWSTR szBuffer, 
                                         DWORD cchBuffer,
                                         DWORD* dwLength);

void InitializeClrVerInfo(ClrVerInfo& clrVerInfo)
{
    clrVerInfo.ClrInstalled = FALSE;
	clrVerInfo.v2PlusInstalled = FALSE;
    clrVerInfo.LatestRuntime[0] = L'\0';
	clrVerInfo.LatestRuntimeNumChars = 0;
}

void GetClrVerInfo(ClrVerInfo& clrVerInfo)
{
    PGetRRI GetRequestedRuntimeInfoFunc = NULL;
    PGetCV GetCorVersionFunc = NULL;
    HMODULE MscoreeHandle = NULL;
    HRESULT hr = S_OK;

    // First, if mscoree.dll is not found on the machine, then there aren't any CLRs on the machine
    MscoreeHandle = LoadLibraryA("mscoree.dll");
    if (MscoreeHandle == NULL)
        return;

    // There were certain OS's that shipped with a "placeholder" mscoree.dll. The existance of this DLL
    // doesn't mean there are CLRs installed on the box.
    //
    // If this mscoree doesn't have an implementation for GetCORVersion, then we know it's one of these
    // placeholder dlls.
    GetCorVersionFunc = (PGetCV)GetProcAddress(MscoreeHandle, "GetCORVersion");
    
    if (GetCorVersionFunc == NULL)
        return;

	//hr = GetCorVersionFunc(	clrVerInfo.LatestRuntime,  
	//						NumItems(clrVerInfo.LatestRuntime),
	//						&clrVerInfo.LatestRuntimeNumChars);

	clrVerInfo.ClrInstalled = TRUE;

	// Now get real runtime info function	
    GetRequestedRuntimeInfoFunc = (PGetRRI)GetProcAddress(MscoreeHandle, "GetRequestedRuntimeInfo");

	// The v2.0 shim allows us to use flags for this function that makes it easier to use. The v1.1 mscoree.dll will
    // not allow us to call this function with 3 NULLs. However, the v2.0 mscoree.dll, along with the RUNTIME_INFO_UPGRADE_VERSION
    // flag, will return us the latest version of the CLR on the machine.
	// (Except it gets overridden by the config file as well...)
    hr = GetRequestedRuntimeInfoFunc(NULL, // pExe
                   NULL, // pwszVersion
                   NULL, // ConfigurationFile
                   0, // startupFlags
                   RUNTIME_INFO_UPGRADE_VERSION|RUNTIME_INFO_DONT_RETURN_DIRECTORY|RUNTIME_INFO_DONT_SHOW_ERROR_DIALOG, // runtimeInfoFlags,
                   NULL, // pDirectory
                   0, // dwDirectory
                   NULL, // dwDirectoryLength
                   clrVerInfo.LatestRuntime, // pVersion
                   NumItems(clrVerInfo.LatestRuntime), // cchBuffer
                   &clrVerInfo.LatestRuntimeNumChars); // dwlength

    // If this fails, then v2.0 of mscoree.dll was not installed on the machine.
    if (SUCCEEDED(hr))
		clrVerInfo.v2PlusInstalled = TRUE;
}
