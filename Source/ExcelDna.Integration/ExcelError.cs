//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Runtime.InteropServices;

namespace ExcelDna.Integration
{
    public enum ExcelError : short
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

        private ExcelMissing()
        {
        }
    }

    public class ExcelEmpty
    {
        public static readonly ExcelEmpty Value = new ExcelEmpty();

        private ExcelEmpty()
        {
        }
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
