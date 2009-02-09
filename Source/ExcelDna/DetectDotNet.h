/*
	CDetectDotNet Version 2.2
	Last updated : Jul 21 2005

	Copyright(C) Nishant Sivakumar 2005. All Rights Reserved.	
	Contact Email : nish@voidnish.com
	URLs : www.voidnish.com, blog.voidnish.com	

	Info :-

	CDetectDotNet is a C++ class that detects whether the .NET Framework
	is present on a machine. If .NET is present, it can retrieve a list
	of .NET versions that are present on that machine.	
*/

#pragma once

#include <Windows.h>

// GvD - Added to fix compiling with VS 2005
#define _CONVERSION_DONT_USE_THREAD_LOCALE

#include <atlconv.h>

#pragma comment(lib,"Version.lib")
#include <strsafe.h>

#include <vector>
#include <string>
using namespace std;


/*
	Definition for GetRequestedRuntimeInfo copy/pasted from
		mscoree.h. We do not want a dependency on mscoree.dll
		and so we use LoadLibrary/GetProcAddress.		
*/
typedef HRESULT (__stdcall *FPGetRequestedRuntimeInfo)(
	LPCWSTR pExe, LPCWSTR pwszVersion, LPCWSTR pConfigurationFile, 
	DWORD startupFlags, DWORD runtimeInfoFlags, 
	LPWSTR pDirectory, DWORD dwDirectory, DWORD *dwDirectoryLength, 
	LPWSTR pVersion, DWORD cchBuffer, DWORD* dwlength);

class CDetectDotNet
{
protected:
	LPTSTR m_szInstallRootPath;
	bool m_bDotNetPresent;
public:
	CDetectDotNet();
	~CDetectDotNet();

	// Returns true if .NET is detected, false otherwise
	bool IsDotNetPresent();

	/* 
		Gets the base path (from the registry) where the various
		versions of the CLR are installed.
	*/
	bool GetInstallRootPath(TCHAR* szRootPath, DWORD dwBufferSize);

	/*
		CLRVersions - This vector<string> will be populated with the various
			CLR versions that are present on the system. 

		Returns - The count of CLR versions detected.
	*/
	size_t EnumerateCLRVersions(vector<string>& CLRVersions);
protected:

	//Detects .NET and caches the result for IsDotNetPresent
	bool IsDotNetPresentInternal();

	//Caches the root path for GetInstallRootPath
	bool GetInstallRootPathInternal(TCHAR* szRootPath, DWORD dwBufferSize);

	/*
		Internal function that detects the possible list of CLR versions.
		It does this by enumerating the .NET root path (older CLR versions
		that have been un-installed leave their directories behind - so this
		actually returns a lot more versions that are actually available).

		Returns - The count of possible CLR versions on the system.
	*/
	size_t EnumeratePossibleCLRVersionsInternal(vector<string>& CLRVersions);

	/*
		Internal function that uses GetRequestedRuntimeInfo to figure
		out if a specific CLR version exists on the system.

		Returns - true if that version is detected, false otherwise. 
	*/
	bool CheckForSpecificCLRVersionInternal(LPCWSTR pwszVersion);	

	/*
		Internal function that extracts the file version from an assembly 
		and returns the .NET version in major.minor.build form.
	*/
	string GetDotNetVersion(LPCTSTR szFolder);

	/*
		Internal function - Given a folder, it'll search for mscorlib.dll 
		within this folder and extract its version.
	*/
	string GetVersionFromFolderName(string szFolderName);
};
