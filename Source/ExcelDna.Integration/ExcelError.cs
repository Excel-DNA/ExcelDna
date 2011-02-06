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

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration
{
	public enum ExcelError : ushort
	{
		ExcelErrorNull = 0,
		ExcelErrorDiv0 = 7,
		ExcelErrorValue = 15,
		ExcelErrorRef = 23,
		ExcelErrorName = 29,
		ExcelErrorNum = 36,
		ExcelErrorNA = 42,
        ExcelErrorGettingData = 43
	}

    public class ExcelMissing
    {
        public static readonly ExcelMissing Value = new ExcelMissing();
        private ExcelMissing() {}
    }

    public class ExcelEmpty
    {
        public static readonly ExcelEmpty Value = new ExcelEmpty();
        private ExcelEmpty() { }
    }
}
