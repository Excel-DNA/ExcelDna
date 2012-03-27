/*
  Copyright (C) 2005-2012 Govert van Drimmelen

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

    public static class ExcelErrorUtil
    {
        public static object ToComError(ExcelError excelError)
        {
            // From this post: http://groups.google.com/group/exceldna/browse_frm/thread/67a9a6c49e0b49b3
            int code;
            switch (excelError)
            {
                case ExcelError.ExcelErrorNull:
                    code = -2146826288;
                    break;
                case ExcelError.ExcelErrorDiv0:
                    code = -2146826281;
                    break;
                case ExcelError.ExcelErrorRef:
                    code = -2146826265;
                    break;
                case ExcelError.ExcelErrorName:
                    code = -2146826259;
                    break;
                case ExcelError.ExcelErrorNum:
                    code = -2146826252;
                    break;
                case ExcelError.ExcelErrorNA:
                    code = -2146826246;
                    break;
                case ExcelError.ExcelErrorGettingData:
                case ExcelError.ExcelErrorValue:
                default:
                    code = -2146826273;
                    break;
            }
            return new ErrorWrapper(code);
        }
    }
}
