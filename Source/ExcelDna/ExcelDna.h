/*
  Copyright (C) 2005-2011 Govert van Drimmelen

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

#pragma once

// EXPORT_COUNT defines the size of the thunk table. 
// Must match the number of functions exported in ExcelDna.cpp and JmpExports64.asm
//#define EXPORT_COUNT 250
#define EXPORT_COUNT 1000

// The function pointers will be exported from managed code in the StdCall convention.
typedef void    (__stdcall *PFN)();
typedef short   (__stdcall *PFN_SHORT_VOID)();
typedef short   (__stdcall *PFN_VOID_LPXLOPER)(void*);
typedef short   (__stdcall *PFN_VOID_LPXLOPER12)(void*);
typedef void*   (__stdcall *PFN_LPXLOPER_LPXLOPER)(void*);
typedef void*   (__stdcall *PFN_LPXLOPER12_LPXLOPER12)(void*);
typedef void    (__stdcall *PFN_PFNEXCEL12)(void*);
typedef HRESULT (__stdcall *PFN_GET_CLASS_OBJECT)(CLSID clsid, IID iid, LPVOID* ppv);
typedef HRESULT (__stdcall *PFN_HRESULT_VOID)();

// ExcelDna add-ins do not implement xlAutoRegister because all
// registrations contain the signature from the start.
struct XlAddInExportInfo
{
	INT32 ExportInfoVersion;
	DWORD AppDomainId;
	PFN_SHORT_VOID				pXlAutoOpen;
	PFN_SHORT_VOID				pXlAutoClose;
	PFN_SHORT_VOID				pXlAutoRemove;
	PFN_VOID_LPXLOPER			pXlAutoFree;
	PFN_VOID_LPXLOPER12			pXlAutoFree12;
	PFN_PFNEXCEL12				pSetExcel12EntryPt;
	PFN_HRESULT_VOID			pDllRegisterServer;
	PFN_HRESULT_VOID			pDllUnregisterServer;
	PFN_GET_CLASS_OBJECT		pDllGetClassObject;
	PFN_HRESULT_VOID			pDllCanUnloadNow;
	// The thunk table that hooks up the fxxx exports from the .xll with the marshaled function pointers.
	INT32  ThunkTableLength;
	PFN*   ThunkTable;           // Actually (PFN ThunkTable[EXPORT_COUNT])
};

XlAddInExportInfo* CreateExportInfo();
