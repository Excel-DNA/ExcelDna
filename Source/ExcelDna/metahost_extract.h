// This file is an extract from the compiler-generated file:
// C++ source equivalent of Win32 type library C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb

#pragma once
#include <comdef.h>

#ifndef __ICLRRuntimeInfo_FWD_DEFINED__
#define __ICLRRuntimeInfo_FWD_DEFINED__
typedef interface ICLRRuntimeInfo ICLRRuntimeInfo;
#endif 	/* __ICLRRuntimeInfo_FWD_DEFINED__ */

extern "C" const GUID __declspec(selectany) IID_ICLRMetaHost = 
	{0xD332DB9E, 0xB9B3, 0x4125, {0x82, 0x07, 0xA1, 0x48, 0x84, 0xF5, 0x32, 0x16}};
extern "C" const GUID __declspec(selectany) CLSID_CLRMetaHost =
	{0x9280188d, 0xe8e, 0x4867, {0xb3, 0xc, 0x7f, 0xa8, 0x38, 0x84, 0xe8, 0xde}};
extern "C" const GUID __declspec(selectany) IID_ICLRRuntimeInfo =
	{0xBD39D1D2, 0xBA2F, 0x486a, {0x89, 0xB0, 0xB4, 0xB0, 0xCB, 0x46, 0x68, 0x91}};


typedef HRESULT ( __stdcall *CallbackThreadSetFnPtr )( void);

typedef HRESULT ( __stdcall *CallbackThreadUnsetFnPtr )( void);

typedef void ( __stdcall *RuntimeLoadedCallbackFnPtr )( 
    ICLRRuntimeInfo* pRuntimeInfo,
    CallbackThreadSetFnPtr pfnCallbackThreadSet,
    CallbackThreadUnsetFnPtr pfnCallbackThreadUnset);

struct __declspec(uuid("D332DB9E-B9B3-4125-8207-A14884F53216"))
ICLRMetaHost : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE GetRuntime( 
        /* [in] */ LPCWSTR pwzVersion,
        /* [in] */ REFIID riid,
        /* [retval][iid_is][out] */ LPVOID *ppRuntime) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE GetVersionFromFile( 
        /* [in] */ LPCWSTR pwzFilePath,
        /* [size_is][out] */ 
        __out_ecount_full(*pcchBuffer)  LPWSTR pwzBuffer,
        /* [out][in] */ DWORD *pcchBuffer) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE EnumerateInstalledRuntimes( 
        /* [retval][out] */ IEnumUnknown **ppEnumerator) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE EnumerateLoadedRuntimes( 
        /* [in] */ HANDLE hndProcess,
        /* [retval][out] */ IEnumUnknown **ppEnumerator) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE RequestRuntimeLoadedNotification( 
        /* [in] */ RuntimeLoadedCallbackFnPtr pCallbackFunction) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE QueryLegacyV2RuntimeBinding( 
        /* [in] */ REFIID riid,
        /* [retval][iid_is][out] */ LPVOID *ppUnk) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE ExitProcess( 
        /* [in] */ INT32 iExitCode) = 0;
        
};

struct __declspec(uuid("BD39D1D2-BA2F-486a-89B0-B4B0CB466891"))
ICLRRuntimeInfo : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE GetVersionString( 
        /* [size_is][out] */ 
        __out_ecount_full_opt(*pcchBuffer)  LPWSTR pwzBuffer,
        /* [out][in] */ DWORD *pcchBuffer) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE GetRuntimeDirectory( 
        /* [size_is][out] */ 
        __out_ecount_full(*pcchBuffer)  LPWSTR pwzBuffer,
        /* [out][in] */ DWORD *pcchBuffer) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE IsLoaded( 
        /* [in] */ HANDLE hndProcess,
        /* [retval][out] */ BOOL *pbLoaded) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE LoadErrorString( 
        /* [in] */ UINT iResourceID,
        /* [size_is][out] */ 
        __out_ecount_full(*pcchBuffer)  LPWSTR pwzBuffer,
        /* [out][in] */ DWORD *pcchBuffer,
        /* [lcid][in] */ LONG iLocaleID) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE LoadLibrary( 
        /* [in] */ LPCWSTR pwzDllName,
        /* [retval][out] */ HMODULE *phndModule) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE GetProcAddress( 
        /* [in] */ LPCSTR pszProcName,
        /* [retval][out] */ LPVOID *ppProc) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE GetInterface( 
        /* [in] */ REFCLSID rclsid,
        /* [in] */ REFIID riid,
        /* [retval][iid_is][out] */ LPVOID *ppUnk) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE IsLoadable( 
        /* [retval][out] */ BOOL *pbLoadable) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE SetDefaultStartupFlags( 
        /* [in] */ DWORD dwStartupFlags,
        /* [in] */ LPCWSTR pwzHostConfigFile) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE GetDefaultStartupFlags( 
        /* [out] */ DWORD *pdwStartupFlags,
        /* [size_is][out] */ 
        __out_ecount_full_opt(*pcchHostConfigFile)  LPWSTR pwzHostConfigFile,
        /* [out][in] */ DWORD *pcchHostConfigFile) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE BindAsLegacyV2Runtime( void) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE IsStarted( 
        /* [out] */ BOOL *pbStarted,
        /* [out] */ DWORD *pdwStartupFlags) = 0;
        
};
