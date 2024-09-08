//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Runtime.InteropServices;

// WARNING: We use IntPtrs for pointers, but often really mean int.

// NOTE: Check http://blogs.msdn.com/b/oldnewthing/archive/2009/08/13/9867383.aspx for 64-bit packing.
//       and http://msdn.microsoft.com/en-us/library/ms973190.aspx for general guidance.

// Regarding unmanaged memory:
// We're currently using the COM Task Memory Allocator. I hoped this would be amenable to cleaning up from the native side too, 
// should we want to. Maybe there are other good options:
// http://www.codeproject.com/Articles/32912/Handling-of-Large-Byte-Arrays
// SafeBuffer class?
// We get error when size is 320000032

// WARNING: The marshalers here are rather particular to the way they are used --
//			mainly to marshal in the _reverse_ direction to what is expected.
//			In particular, this means I allocate native memory only for return parameters
//			and generally only keep one allocation per marshaller class.
//			If that class were used for function parameters in an outgoing call,
//			multiple memory allocations would overwrite each other!
//			For this case there is the Cleanup stuff.
//			The only exception to how I use this is for the object and object[] marshalling
//			in the Excel4v function.

// TODO: Marshalers should implement disposable pattern.

// Revisit marshaling, taking in consideration stackalloc and/or using a static ThreadLocal with a fixed size.
namespace ExcelDna.Loader
{
    // Internal Implementations of the Excel Types
    // CONSIDER: How to (if?) make these available to the user code
    // For now I think of this as an internal structure used in the marshaling

    [StructLayout(LayoutKind.Explicit)]
    unsafe struct XlString12
    {
        [FieldOffset(0)]
        public ushort Length;
        [FieldOffset(2)]
        public fixed char Data[1]; // Actually Data[Length]

        public static readonly int MaxLength = 32767; // chars
    }
    [Flags]
    enum XlType12 : uint
    {
        XlTypeNumber = 0x0001,
        XlTypeString = 0x0002,
        XlTypeBoolean = 0x0004,
        XlTypeReference = 0x0008,
        XlTypeError = 0x0010,
        XlTypeFlow = 0x0020, // Unused
        XlTypeArray = 0x0040,
        XlTypeMissing = 0x0080,
        XlTypeEmpty = 0x0100,
        XlTypeSReference = 0x0400,
        XlTypeInt = 0x0800,     // int16 in XlOper, int32 in XlOper12, never passed into UDF

        XlBitXLFree = 0x1000,
        XlBitDLLFree = 0x4000,

        XlTypeBigData = XlTypeString | XlTypeInt    // Used only for marshaling the async handle for Excel 2010+ native async.
                                                    // Sometimes used by Excel to return a pointer
                                                    // For example in xlGetInstPtr: http://msdn.microsoft.com/en-us/library/ff475872.aspx
    }

    [StructLayout(LayoutKind.Explicit)]
    unsafe struct XlOper12
    {
        [StructLayout(LayoutKind.Sequential)]
        unsafe public struct XlOper12Array
        {
            public XlOper12* pOpers;
            public int Rows;
            public int Columns;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct XlRectangle12
        {
            [FieldOffset(0)]
            public int RowFirst;
            [FieldOffset(4)]
            public int RowLast;
            [FieldOffset(8)]
            public int ColumnFirst;
            [FieldOffset(12)]
            public int ColumnLast;

            public XlRectangle12(int rowFirst, int rowLast, int columnFirst, int columnLast)
            {
                // CONSIDER: Range checking.
                RowFirst = rowFirst;
                RowLast = rowLast;
                ColumnFirst = columnFirst;
                ColumnLast = columnLast;
            }
        }

        [StructLayout(LayoutKind.Explicit)]
        unsafe public struct XlMultiRef12
        {
            [FieldOffset(0)]
            public ushort Count;
            [FieldOffset(4)]
            public XlRectangle12 Rectangles;    // Not XlRectangle12, actually Rectangles[Count] !
        }

        [StructLayout(LayoutKind.Explicit)]
        unsafe public struct XlSReference12
        {
            [FieldOffset(0)]
            public ushort Count;                // Always = 1
            [FieldOffset(4)]
            public XlRectangle12 Reference;
        }

        [StructLayout(LayoutKind.Sequential)]
        unsafe public struct XlReference12
        {
            public XlMultiRef12* pMultiRef;
            public IntPtr SheetId;
        }

        [StructLayout(LayoutKind.Sequential)]
        unsafe public struct XlBigData
        {
            public IntPtr hData;    // Handle or byte*, but we only use it as a handle.
            public long cbData;
        }

        [FieldOffset(0)]
        public double numValue;
        [FieldOffset(0)]
        public XlString12* pstrValue;
        [FieldOffset(0)]
        public int boolValue;
        [FieldOffset(0)]
        public int intValue;
        [FieldOffset(0)]
        public int /*ExcelError*/ errValue;
        [FieldOffset(0)]
        public XlOper12Array arrayValue;
        [FieldOffset(0)]
        public XlReference12 refValue;
        [FieldOffset(0)]
        public XlSReference12 srefValue;
        [FieldOffset(0)]
        public XlBigData bigData;
        [FieldOffset(24)]
        public XlType12 xlType;
    }
    // Excel signature type 'K'
    /* From Excel97DevKit:
	 * K Data Type
	 * The K data type uses a pointer to a variable-size FP structure. 
	 * You should define this structure in the DLL or code resource as follows:

		typedef struct _FP12
		{
			INT32 rows;
			INT32 columns;
			double array[1];        // Actually, array[rows][columns]
		} FP12;

	 *	The declaration double array[1] allocates storage only for a single-element array. 
	 *  The number of elements in the actual array equals the number of rows multiplied 
	 *  by the number of columns.

	 */
    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    unsafe struct XlFP12
    {
        public int Rows;
        public int Columns;
        public fixed double Values[1]; // Actually, Values[rows][columns]
    }

    static class XlTypes
    {
        // These are identifiers for xlfRegister types used in the pointer-only direct marshalling 
        public const string Xloper = "Q";
        public const string XloperAllowRef = "U";
        public const string DoublePtr = "E";        // double*
        public const string String = "D%";          // XLSTRING12
        public const string DoubleArray = "K%";     // FP12*
        public const string BoolPtr = "L";          // short*
        public const string AsyncHandle = "X";
    }

}
