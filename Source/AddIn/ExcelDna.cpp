/*
  Copyright (C) 2005, 2006 Govert van Drimmelen

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

// This file is compiled as managed code - It Just Works!

#include "JumpTable.h"

#using "ExcelDna.Integration.netmodule"

// Managed namespaces
using namespace System;
using namespace ExcelDna::Integration;

// The callback type exported to DnaXll to allow it to set the jumps
void SetJumpManaged(int fi, IntPtr pfn)
{
	SetJump(fi, (PFN)pfn.ToPointer());
}
	
// Actual Excel Add-In API exports
extern "C"
{
	__declspec(dllexport) short xlAutoOpen()
	{
		XlLibrary::SetJump = gcnew SetJumpDelegate(SetJumpManaged);
		return XlLibrary::AutoOpen();
	}

	__declspec(dllexport) short xlAutoClose()
	{
		return XlLibrary::AutoClose();
	}

	__declspec(dllexport) short xlAutoAdd()
	{
		return XlLibrary::AutoAdd();
	}

	__declspec(dllexport) short xlAutoRemove()
	{
		return XlLibrary::AutoRemove();
	}

	__declspec(dllexport) void xlAutoFree(void* pXloper)
	{
		return XlLibrary::AutoFree((IntPtr)pXloper);
	}

	__declspec(dllexport) void* xlAddInManagerInfo(void* pXloper)
	{
		return (void*)XlLibrary::AddInManagerInfo((IntPtr)pXloper);
	}
}

