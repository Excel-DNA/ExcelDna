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

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

// WARNING: We use IntPtrs for pointers, but often really mean int.
// On 64-bit platform, I don't know what would be appropriate. Probably won't work. 
// Packing and the like might be very different! 

// WARNING: The marshalers here are rather particular to the way they are used --
//			mainly to marshal in the _reverse_ direction to what is expected.
//			In particular, this means I allocate native memory only for return parameters
//			and generally only keep one allocation per marshaller class.
//			If that class were used for function parameters in an outgoing call,
//			multiple memory allocations would overwrite each other!
//			For this case there is the Cleanup stuff.
//			The only exception to how I use this is for the object and object[] marshalling
//			in the Excel4v function.

// TODO: Check what happens for re-entrancy, e.g. Calling a UDF from Excel.Excel4 !!

// TODO: Marshalers should implement disposable pattern.

namespace ExcelDna.Loader
{
	// Internal Implementations of the Excel Types
	// CONSIDER: How to (if?) make these available to the user code
	// For now I think of this as an internal structure used in the marshaling

    [Flags]
    internal enum XlType : ushort
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

        XlBitXLFree = 0x1000,	// Unused so far
        XlBitDLLFree = 0x4000,

        XlTypeBigData = XlTypeString | XlTypeInt	// Unused so far (IntPtr)
    }

    [StructLayout(LayoutKind.Explicit)]
    internal unsafe struct XlString
    {
        [FieldOffset(0)]
        public byte Length;
        [FieldOffset(1)]
        public fixed byte Data[255];
    }

	[StructLayout(LayoutKind.Explicit)]
	internal unsafe struct XlOper
	{
        [StructLayout(LayoutKind.Explicit)]
		unsafe public struct XlOperArray
		{
			[FieldOffset(0)]
			public XlOper* pOpers;
			[FieldOffset(4)]
			public ushort Rows;
			[FieldOffset(6)]
			public ushort Columns;
		}

        [StructLayout(LayoutKind.Explicit)]
        public struct XlRectangle
        {
            [FieldOffset(0)]
            public ushort RowFirst;
            [FieldOffset(2)]
            public ushort RowLast;
            [FieldOffset(4)]
            public byte ColumnFirst;
            [FieldOffset(5)]
            public byte ColumnLast;

            /*
            public XlRectangle(ushort rowFirst, ushort rowLast, byte columnFirst, byte columnLast)
            {
                RowFirst = rowFirst;
                RowLast = rowLast;
                ColumnFirst = columnFirst;
                ColumnLast = columnLast;
            }
             * */

            // DOCUMENT: if the values for row / column exceed the limits,
            // they are silently trimmed to valid values.
            // This keeps the external interface CLS-compliant and 
            // will assist future compatibility when the sheets get bigger in Excel12
            public XlRectangle(int rowFirst, int rowLast, int columnFirst, int columnLast)
            {
                RowFirst = (ushort)Math.Max(0, Math.Min(rowFirst, ushort.MaxValue));
                RowLast = (ushort)Math.Max(0, Math.Min(rowLast, ushort.MaxValue));
                ColumnFirst = (byte)Math.Max(0, Math.Min(columnFirst, byte.MaxValue)); ;
                ColumnLast = (byte)Math.Max(0, Math.Min(columnLast, byte.MaxValue)); ; ;
            }
        }

		[StructLayout(LayoutKind.Explicit)]
		unsafe public struct XlMultiRef
		{
			[FieldOffset(0)]
			public ushort Count;
			[FieldOffset(2)]
			public XlRectangle Rectangles;	// Not XlRectangle, actually XlRectangle[Count] !
		}


		[StructLayout(LayoutKind.Explicit)]
		unsafe public struct XlSReference
		{
			[FieldOffset(0)]
			public ushort Count;
			[FieldOffset(2)]
			public XlRectangle Reference;
		}


		[StructLayout(LayoutKind.Sequential)]
		unsafe public struct XlReference
		{
			public XlMultiRef* pMultiRef;
			public IntPtr SheetId;

		}

		[FieldOffset(0)]
		public double numValue;
		[FieldOffset(0)]
		public XlString* pstrValue;
		[FieldOffset(0)]
		public ushort boolValue;
		[FieldOffset(0)]
		public short intValue;
		[FieldOffset(0)]
		public ushort /*ExcelError*/ errValue;
		[FieldOffset(0)]
		public XlOperArray arrayValue;
		[FieldOffset(0)]
		public XlReference refValue;
		[FieldOffset(0)]
		public XlSReference srefValue;
		[FieldOffset(8)]
		public XlType xlType;
	}



	// DOCUMENT: Returned strings longer than 255 chars are truncated.
	public class XlStringReturnMarshaler : ICustomMarshaler
	{

		static ICustomMarshaler instance;
		IntPtr pNative; // Pointer to XlString, allocated once on initialization

		public XlStringReturnMarshaler()
		{
			int size = Marshal.SizeOf(typeof(XlString));
			pNative = Marshal.AllocCoTaskMem(size);
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			if (instance == null)
				instance = new XlStringReturnMarshaler();
			return instance;
		}

		unsafe public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			// CONSIDER: Checking for null, checking object type.
			// CONSIDER: Marshal back as OPER for errors etc.

			// DOCUMENT: This function is not called if the return is null!
			// DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!

			String str = (String)ManagedObj;
			XlString* pdest = (XlString*)pNative;
			int charCount = Math.Min(str.Length, 255);
			fixed (char* psrc = str )
			{
                // Support for system codepage by hmd
				//int written = Encoding.ASCII.GetBytes(psrc, charCount, pdest->Data, 255);
                Encoding enc = Encoding.GetEncoding(ASCIIEncoding.Default.CodePage,
                                                    EncoderFallback.ReplacementFallback,    
                                                    DecoderFallback.ReplacementFallback);
                int written = enc.GetBytes(psrc, charCount, pdest->Data, 255);
				pdest->Length = (byte)written;
			}
			
			return pNative;
		}

		public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			throw new NotImplementedException("This marshaler only used for managed to native return type marshaling.");
		}

		public void CleanUpManagedData(object ManagedObj) { }
		public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
		public int GetNativeDataSize() { return -1; }
	}

	// Boolean returns are returned as an XLOPER 
	// - can't make it short due to marshaling limitations,
	// so we force a boxing
	public unsafe class XlBooleanReturnMarshaler : ICustomMarshaler
	{
		static ICustomMarshaler instance;
		IntPtr pNative; // this is really an XlOper, and is is allocated once, 
						// when the marshaller is constructed, 
						// and is never reclaimed

		public XlBooleanReturnMarshaler()
		{
			int size = Marshal.SizeOf(typeof(XlOper));
			pNative = Marshal.AllocCoTaskMem(size);
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			if (instance == null)
				instance = new XlBooleanReturnMarshaler();
			return instance;

		}

		public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			XlOper* xlOper = (XlOper*)pNative;
			xlOper->boolValue = (bool)ManagedObj ? (ushort)1 : (ushort)0;
			xlOper->xlType = XlType.XlTypeBoolean;
			return pNative;
		}

		public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			throw new NotImplementedException("This marshaler only used for managed to native return type marshaling.");
		}

		public void CleanUpManagedData(object ManagedObj) { }
		public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
		public int GetNativeDataSize() { return -1; }
	}

	// DateTimes are returned as a double*
    public class XlDateTimeMarshaler : ICustomMarshaler
	{
		static ICustomMarshaler instance;
		IntPtr pNative;	// points to a double - never changes

		public XlDateTimeMarshaler()
		{
			int size = Marshal.SizeOf(typeof(double));
			pNative = Marshal.AllocCoTaskMem(size);
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			if (instance == null)
				instance = new XlDateTimeMarshaler();
			return instance;
		}

		unsafe public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			*(double*)pNative = ((DateTime)ManagedObj).ToOADate();
			return pNative;
		}

		unsafe public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			try
			{
				double dateSerial = *(double*)pNativeData;
				return DateTime.FromOADate(dateSerial);
			}
			catch
			{
				// This case is where the range of the OADate is exceeded.
				// By returning null, the unboxing code will fail,
				// causing a runtime exception that is caught and returned as a #Value error.
				return null;
			}
		}

		public void CleanUpManagedData(object ManagedObj) { }
		public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
		public int GetNativeDataSize() { return -1; }
	}

	// Excel signature type 'K'
	/* From Excel97DevKit:
	 * K Data Type
	 * The K data type uses a pointer to a variable-size FP structure. 
	 * You should define this structure in the DLL or code resource as follows:

		typedef struct _FP
		{
			unsigned short int rows;
			unsigned short int columns;
			double array[1];        // Actually, array[rows][columns]
		} FP;

	 *	The declaration double array[1] allocates storage only for a single-element array. 
	 *  The number of elements in the actual array equals the number of rows multiplied 
	 *  by the number of columns.

	 */
	[StructLayout(LayoutKind.Sequential, Pack = 8)]
	internal unsafe struct XlFP
	{
		public ushort Rows;
		public ushort Columns;
		public fixed double Values[1];
	}

    public class XlDoubleArrayMarshaler : ICustomMarshaler
	{

		// CONSIDER: Marshal all return types of double[,] as XLOPER
		// and set xlFree bit, and handle callback.
		// This will reduce memory usage but be slower, as we would get callback
		// into managed code, unless we implement xlFree in native
		// (we can use Com memory allocator to free there)
		// For now just do fast, simple, slightly memory hogging thing.

        static XlDoubleArrayMarshaler instance1;	// For rank 1 arrays
        static XlDoubleArrayMarshaler instance2;	// For rank 2 arrays

		int rank;
		IntPtr pNative; // For managed -> native returns 
						// This points to the last FP that was marshaled.
						// FPs are re-allocated on every managed->native transition

		public XlDoubleArrayMarshaler(int rank)
		{
			this.rank = rank;
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			// marshalCookie denotes the array rank
			// must be 1 or 2
			if (marshalCookie == "1")
			{
				if (instance1 == null)
					instance1 = new XlDoubleArrayMarshaler(1);
				return instance1;
			}
			else if (marshalCookie == "2")
			{
				if (instance2 == null)
					instance2 = new XlDoubleArrayMarshaler(2);
				return instance2;
			}
			throw new ArgumentException("Invalid cookie for XlDoubleArrayMarshaler");
		}

		unsafe public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			// CONSIDER: Checking for null, checking object type
			// CONSIDER: Managing memory differently
			// Here we allocate and clear when the next array is returned
			// we might also return XLOPER and have xlFree called back.

			// If array is too big!?, we just truncate

			// TODO: Remove duplication - due to fixed / pointer interaction

			// DOCUMENT: This function is not called if the return is null!
			// DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!

			Marshal.FreeCoTaskMem(pNative);
			pNative = IntPtr.Zero;
			
			ushort rows;
			ushort columns;
			int allColumns;	// all in the managed array
			if (rank == 1)
			{
				double[] doubles = (double[])ManagedObj;

				rows = 1;
				allColumns = doubles.Length;
				columns = (ushort)Math.Min(allColumns, ushort.MaxValue);

                // Guard against invalid arrays - with no columns.
                // Just return null, which Excel will turn into #NUM
                if (columns == 0)
                    return IntPtr.Zero;

				fixed(double* src = doubles)
				{
					AllocateFPAndCopy(src, rows, columns, allColumns);	
				}
			}
			else if (rank == 2)
			{
				double[,] doubles = (double[,])ManagedObj;

				rows = (ushort)Math.Min(doubles.GetLength(0), ushort.MaxValue);
				allColumns = doubles.GetLength(1);
				columns = (ushort)Math.Min(allColumns, ushort.MaxValue);

                // Guard against invalid arrays - with no rows or no columns.
                // Just return null, which Excel will turn into #NUM
                if (rows == 0 || columns == 0)
                    return IntPtr.Zero;

				fixed (double* src = doubles)
				{
					AllocateFPAndCopy(src, rows, columns, allColumns);
				}
			}
			else
			{
				throw new InvalidOperationException("Damaged XlDoubleArrayMarshaler rank");
			}

			return pNative;
		}

		unsafe private void AllocateFPAndCopy(double* pSrc, ushort rows, ushort columns, int allColumns)
		{
            XlFP* pFP;
			int size = Marshal.SizeOf(typeof(XlFP)) +
			   Marshal.SizeOf(typeof(double)) * (rows * columns - 1); // room for one double is already in FP struct
			pNative = Marshal.AllocCoTaskMem(size);

			pFP = (XlFP*)pNative;
			pFP->Rows = rows;
			pFP->Columns = columns;
			int count = rows * columns;
			if (columns == allColumns)
			{
				// Fast copy
				CopyDoubles(pSrc, pFP->Values, count);				
			}
			else
			{
				for (int i = 0; i < rows; i++)
				{
					for (int j = 0; j < columns; j++)
					{
						pFP->Values[i] = pSrc[i * allColumns + j];
					}
				}
			}
		}

		unsafe public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			object result;
			XlFP* pFP = (XlFP*)pNativeData;
			
			// Duplication here, because the types are different and wrapped in fixed blocks
			if (rank == 1)
			{
				double[] array;
                if (pFP->Columns == 1)
                {
                    // Take the one and only column as the array
                    array = new double[pFP->Rows];
                }
                else
                {
                    // Take only the first row of the array.
                    array = new double[pFP->Columns];
                }
                // Copy works for either case, due to in-memory layout!
                fixed (double* dest = array)
                {
                    CopyDoubles(pFP->Values, dest, array.Length);
                }
				result = array;
			}
			else if (rank == 2)
			{
				double[,] array = new double[pFP->Rows, pFP->Columns];
				fixed (double* dest = array)
				{
					CopyDoubles(pFP->Values, dest, array.Length);
				}
				result = array;
			}
			else
			{
				Debug.Fail("Damaged XlDoubleArrayMarshaler rank");
				throw new InvalidOperationException("Damaged XlDoubleArrayMarshaler rank");
			}
			return result;
		}

		unsafe private void CopyDoubles(double* pSrc, double* pDest, int count)
		{
			for (int i = 0; i < count; i++)
			{
				pDest[i] = pSrc[i];
			}
		}

		public void CleanUpManagedData(object ManagedObj) {}
		public void CleanUpNativeData(IntPtr pNativeData) {} // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
		public int GetNativeDataSize() { return -1; }
	}

    public class XlObjectMarshaler : ICustomMarshaler
	{
        // Shared instance used for all return values
		static ICustomMarshaler instance;

		IntPtr pNative; // this is really an XlOper, and it is allocated once, 
                        // when the marshaller is constructed.

		public XlObjectMarshaler()
		{
			int size = Marshal.SizeOf(typeof(XlOper));
			pNative = Marshal.AllocCoTaskMem(size);
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			if (instance == null)
				instance = new XlObjectMarshaler();
			return instance;
		}

		unsafe public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			// CONSIDER: Managing memory differently
			// Here we allocate and clear when the next array is returned
			// we might also return XLOPER with the right bits set and have xlFree called back.

			// DOCUMENT: This function is not called if the return is null!
			// (A null pointer is immediately returned to Excel, resulting in #NUM!)
			// However, we allow null for the Excel4 marshaling case, and create an Empty xlOper
            // ^^^^^^^^^^^^ - null case for Excel4 has now been removed ????
			if (ManagedObj is double)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->numValue = (double)ManagedObj;
				pOper->xlType = XlType.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is string)
			{
				// TODO: Consolidate these?
				ICustomMarshaler m = XlStringReturnMarshaler.GetInstance("");
				XlOper* pOper = (XlOper*)pNative;
				pOper->pstrValue = (XlString*)m.MarshalManagedToNative(ManagedObj);
				pOper->xlType = XlType.XlTypeString;
				return pNative;
			}
			else if (ManagedObj is DateTime)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->numValue = ((DateTime)ManagedObj).ToOADate();
				pOper->xlType = XlType.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is bool)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->boolValue = (bool)ManagedObj ? (ushort)1 : (ushort)0;
				pOper->xlType = XlType.XlTypeBoolean;
				return pNative;
			}
			else if (ManagedObj is object[])
			{
				// Redirect to the ObjectArray Marshaler
				// CONSIDER: This might cause some memory to get stuck, 
				// since the memory for the array marshaler is not the same as for this
				ICustomMarshaler m = XlObjectArrayMarshaler.GetInstance("1");
				return m.MarshalManagedToNative(ManagedObj);
			}
			else if (ManagedObj is object[,])
			{
				// Redirect to the ObjectArray Marshaler
				// CONSIDER: This might cause some memory to get stuck, 
				// since the memory for the array marshaler is not the same as for this
				ICustomMarshaler m = XlObjectArrayMarshaler.GetInstance("2");
				return m.MarshalManagedToNative(ManagedObj);
			}
			else if (ManagedObj is double[])
			{
				double[] doubles = (double[])ManagedObj;
				object[] objects = new object[doubles.Length];
				Array.Copy(doubles, objects, doubles.Length);
				ICustomMarshaler m = XlObjectArrayMarshaler.GetInstance("1");
				return m.MarshalManagedToNative(objects);
			}
			else if (ManagedObj is double[,])
			{
				double[,] doubles = (double[,])ManagedObj;
				object[,] objects = new object[doubles.GetLength(0), doubles.GetLength(1)];
				Array.Copy(doubles, objects, doubles.GetLength(0) * doubles.GetLength(1));
				ICustomMarshaler m = XlObjectArrayMarshaler.GetInstance("2");
				return m.MarshalManagedToNative(objects);
			}
			else if (IntegrationMarshalHelpers.IsExcelErrorObject(ManagedObj))
			{
				XlOper* pOper = (XlOper*)pNative;
                pOper->errValue = (ushort)IntegrationMarshalHelpers.ExcelErrorGetValue(ManagedObj);
				pOper->xlType = XlType.XlTypeError;
				return pNative;
			}
            else if (IntegrationMarshalHelpers.IsExcelMissingObject(ManagedObj))
            {
                XlOper* pOper = (XlOper*)pNative;
                pOper->xlType = XlType.XlTypeMissing;
                return pNative;
            }
            else if (IntegrationMarshalHelpers.IsExcelEmptyObject(ManagedObj))
            {
                XlOper* pOper = (XlOper*)pNative;
                pOper->xlType = XlType.XlTypeEmpty;
                return pNative;
            }
            // 13 November -- this should never be called...
                // since this marshaler is not used for the Excel4 call anymore ...?
            //else if (ManagedObj is ExcelReference)
            //{
            //    XlOper* pOper = (XlOper*)pNative;
            //    ExcelReference r = (ExcelReference)ManagedObj;
            //    int refCount = r.InnerReferences.Count;
            //    int numBytes = Marshal.SizeOf(typeof(ushort)) + refCount * Marshal.SizeOf(typeof(XlOper.XlRectangle));
            //    XlOper.XlMultiRef* pMultiRef = (XlOper.XlMultiRef*)Marshal.AllocCoTaskMem(numBytes);
            //    // 13 November 2006 -- Where is the just allocated memory freed?
            //    XlOper.XlReference.SetExcelReference(pOper, pMultiRef, r);
            //    return pNative;
            //}
			else if (ManagedObj is short)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->numValue = (double)((short)ManagedObj);
				pOper->xlType = XlType.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is Missing)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->xlType = XlType.XlTypeMissing;
				return pNative;
			}
			else if (ManagedObj is int)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->numValue = (double)((int)ManagedObj);
				pOper->xlType = XlType.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is uint)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->numValue = (double)((uint)ManagedObj);
				pOper->xlType = XlType.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is byte)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->numValue = (double)((byte)ManagedObj);
				pOper->xlType = XlType.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is ushort)
			{
				XlOper* pOper = (XlOper*)pNative;
				pOper->numValue = (double)((ushort)ManagedObj);
				pOper->xlType = XlType.XlTypeNumber;
				return pNative;
			}
            else if (ManagedObj is decimal)
            {
                XlOper* pOper = (XlOper*)pNative;
                pOper->numValue = (double)((decimal)ManagedObj);
                pOper->xlType = XlType.XlTypeNumber;
                return pNative;
            }
            else if (ManagedObj is float)
            {
                XlOper* pOper = (XlOper*)pNative;
                pOper->numValue = (double)((float)ManagedObj);
                pOper->xlType = XlType.XlTypeNumber;
                return pNative;
            }
            else if (ManagedObj is long)
            {
                XlOper* pOper = (XlOper*)pNative;
                pOper->numValue = (double)((long)ManagedObj);
                pOper->xlType = XlType.XlTypeNumber;
                return pNative;
            }
            else if (ManagedObj is ulong)
            {
                XlOper* pOper = (XlOper*)pNative;
                pOper->numValue = (double)((ulong)ManagedObj);
                pOper->xlType = XlType.XlTypeNumber;
                return pNative;
            }
            // 13 November 2006 -- this marshaler is not used to set up the Excel4 return value anymore.
            //else if (ManagedObj == null)
            //{
            //    // This is never the case for regular marshaling, only for 
            //    // return value for Excel4 function
            //    XlOper* pOper = (XlOper*)pNative;
            //    pOper->xlType = XlType.XlTypeEmpty;
            //    return pNative;
            //}
            else
            {
                // Default error return
                XlOper* pOper = (XlOper*)pNative;
                pOper->errValue = (ushort)IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                pOper->xlType = XlType.XlTypeError;
                return pNative;
            }
		}

		unsafe public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			// Make a nice object from the native OPER
            if (pNativeData == IntPtr.Zero)
            {
                // We don't expect this at all.
                return IntegrationMarshalHelpers.GetExcelEmptyValue();
            }

            object managed;
            XlOper* pOper = (XlOper*)pNativeData;
			// Ignore any Free flags
			XlType type = pOper->xlType & ~XlType.XlBitXLFree & ~XlType.XlBitDLLFree;
			switch (type)
			{
				case XlType.XlTypeNumber:
					managed = pOper->numValue;
					break;
				case XlType.XlTypeString:
                    XlString* pString = pOper->pstrValue;
                    // Support for system codepage by hmd
                    // managed = new string((sbyte*)pString->Data, 0, pString->Length, Encoding.ASCII);
                    Encoding enc = Encoding.GetEncoding(ASCIIEncoding.Default.CodePage,
                                                        EncoderFallback.ReplacementFallback,
                                                        DecoderFallback.ReplacementFallback);
                    managed = new string((sbyte*)pString->Data, 0, pString->Length, enc);
					break;
				case XlType.XlTypeBoolean:
					managed = pOper->boolValue == 1;
					break;
				case XlType.XlTypeError:
					managed = IntegrationMarshalHelpers.GetExcelErrorObject(pOper->errValue);
					break;
				case XlType.XlTypeMissing:
                    // DOCUMENT: Changed in version 0.17.
					// managed = System.Reflection.Missing.Value;
                    managed = IntegrationMarshalHelpers.GetExcelMissingValue();
					break;
				case XlType.XlTypeEmpty:
                    // DOCUMENT: Changed in version 0.17.
                    // managed = null;
                    managed = IntegrationMarshalHelpers.GetExcelEmptyValue();
					break;
				case XlType.XlTypeArray:
					int rows = pOper->arrayValue.Rows;
					int columns = pOper->arrayValue.Columns;
					object[,] array = new object[rows, columns];
					XlOper* opers = (XlOper*)pOper->arrayValue.pOpers;
					for (int i = 0; i < rows; i++)
					{
						for (int j = 0; j < columns; j++)
						{
							int pos = i * columns + j;
							array[i, j] = MarshalNativeToManaged((IntPtr)(opers + pos));
						}
					}
					managed = array;
					break;
				case XlType.XlTypeInt:
					managed = (double)pOper->intValue; // int16 in XlOper // always return double
					break;
				case XlType.XlTypeReference:
					object /*ExcelReference*/ r;
					if (pOper->refValue.pMultiRef == (XlOper.XlMultiRef*)IntPtr.Zero)
					{
						r = IntegrationMarshalHelpers.CreateExcelReference(0, 0, 0, 0, pOper->refValue.SheetId);
					}
					else
					{
						ushort numAreas = *(ushort*)pOper->refValue.pMultiRef;
						XlOper.XlRectangle* pAreas = (XlOper.XlRectangle*)((uint)pOper->refValue.pMultiRef + sizeof(ushort));
                        r = IntegrationMarshalHelpers.CreateExcelReference(
							pAreas[0].RowFirst, pAreas[0].RowLast,
							pAreas[0].ColumnFirst, pAreas[0].ColumnLast, pOper->refValue.SheetId);
						for (int i = 1; i < numAreas; i++)
						{
                            IntegrationMarshalHelpers.ExcelReferenceAddReference(r, 
							               pAreas[i].RowFirst, pAreas[i].RowLast,
										   pAreas[i].ColumnFirst, pAreas[i].ColumnLast);
						}
					}
					managed = r;
					break;
				case XlType.XlTypeSReference:
                    IntPtr sheetId = XlCallImpl.GetCurrentSheetId4();
					object /*ExcelReference*/ sref;
					sref = IntegrationMarshalHelpers.CreateExcelReference(
                                            pOper->srefValue.Reference.RowFirst,
											pOper->srefValue.Reference.RowLast,
											pOper->srefValue.Reference.ColumnFirst,
											pOper->srefValue.Reference.ColumnLast, 
											sheetId /*Current sheet (not Active sheet!)*/);
					managed = sref;
					break;
                case 0:
                    // We get type == 0 when a long (>255 char) string is embedded in an array.
                    // To be consistent with the string handling, we set the value to #VALUE
                    managed = IntegrationMarshalHelpers.GetExcelErrorObject(15 /* ExcelErrorValue */);
                    break;
				default:
                    // Unexpected !? (BigData perhaps - How did it get here?)
                    // We do #VALUE here too, rather than set to null.
                    managed = IntegrationMarshalHelpers.GetExcelErrorObject(15 /* ExcelErrorValue */);
					break;
			}
			return managed;
		}

		public void CleanUpManagedData(object ManagedObj) { }
        public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
		public int GetNativeDataSize() { return -1; }
	}

    public class XlObjectArrayMarshaler : ICustomMarshaler, IDisposable
	{
		// CONSIDER: Marshal return types of object[,] as XLOPER
		// and set xlFree bit, and handle callback.
		// This will reduce memory usage but be slower, as we would get callback
		// into managed code, unless we implement xlFree in native
		// (we can use Com memory allocator to free there)
		// For now just do fast, simple, slightly memory hogging thing.

        static XlObjectArrayMarshaler instance1;	// For rank 1 arrays
        static XlObjectArrayMarshaler instance2;	// For rank 2 arrays

        int rank;
        List<XlObjectArrayMarshaler> nestedInstances = new List<XlObjectArrayMarshaler>();

		bool isExcel4v;	// Used for calls to Excel4 -- flags that returned native data should look different

		IntPtr pNative; // For managed -> native returns 
		// This points to the last OPER (and contained OPER array) that was marshaled
		// OPERs are re-allocated on every managed->native transition
		IntPtr pNativeStrings;
		IntPtr pNativeReferences;
		
        IntPtr pOperPointers; // Used for calls to Excel4v - points to the array of oper addresses

        public XlObjectArrayMarshaler()
        {
            this.rank = 0;  // Must be set before use.
            this.isExcel4v = false;
        }

		public XlObjectArrayMarshaler(int rank)
		{
            this.rank = rank;
            this.isExcel4v = false;
		}

        public XlObjectArrayMarshaler(int rank, bool isExcel4v)
        {
            this.rank = rank;
            this.isExcel4v = isExcel4v;
        }

		public static void FreeMemory()
		{
            // This method is only called via AutoFree for an instance 
            instance1.Reset(true);
            instance2.Reset(true);
		}

        public static ICustomMarshaler GetInstance(string marshalCookie)
        {
            // marshalCookie denotes the array rank
            // must be 1 or 2
            if (marshalCookie == "1")
            {
                if (instance1 == null)
                    instance1 = new XlObjectArrayMarshaler(1);
                return instance1;
            }
            else if (marshalCookie == "2")
            {
                if (instance2 == null)
                    instance2 = new XlObjectArrayMarshaler(2);
                return instance2;
            }
            throw new ArgumentException("Invalid cookie for XlObjectArrayMarshaler");
        }

		unsafe public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			// CONSIDER: Checking for null, checking object type
			// DOCUMENT: This function is not called if the return is null!
			// DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!

			// CONSIDER: Managing memory differently
			// Here we allocate and clear when the next array is returned
			// we might also return XLOPER and have xlFree called back.

			// If array is too big!?, we just truncate

			// TODO: Remove duplication - due to fixed / pointer interaction

			Reset(true);

			ushort rows;
            int rowBase;
			ushort columns; // those in the returned array
            int columnBase;
			int allColumns;	// all in the managed array
			if (rank == 1)
			{
				object[] objects = (object[])ManagedObj;

				rows = 1;
                rowBase = 0;
				allColumns = objects.Length;
				columns = (ushort)Math.Min(objects.Length, ushort.MaxValue);
                columnBase = objects.GetLowerBound(0);
			}
			else if (rank == 2)
			{
				object[,] objects = (object[,])ManagedObj;

				rows = (ushort)Math.Min(objects.GetLength(0), ushort.MaxValue);
                rowBase = objects.GetLowerBound(0);
				allColumns = objects.GetLength(1);
				columns = (ushort)Math.Min(objects.GetLength(1), ushort.MaxValue);
                columnBase = objects.GetLowerBound(1);
			}
			else
			{
				throw new InvalidOperationException("Damaged XlObjectArrayMarshaler rank");
			}

            // Some counters for the multi-pass
            int cbNativeStrings = 0;
			int numReferenceOpers = 0;
			int numReferences = 0;

			// Allocate native space
			int cbNative = Marshal.SizeOf(typeof(XlOper)) +				// OPER that is returned
				Marshal.SizeOf(typeof(XlOper)) * (rows * columns);	// Array of OPER inside the result
			pNative = Marshal.AllocCoTaskMem(cbNative);

			// Set up returned OPER
			XlOper* pOper = (XlOper*)pNative;
            // Excel chokes badly on empty arrays (e.g. crash in function wizard) - rather return the default erro value, #VALUE!
            if (rows * columns == 0)
            {
                pOper->errValue = (ushort)IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                pOper->xlType = XlType.XlTypeError;
            }
            else
            {
                // Some contents to put into an array....
                pOper->xlType = XlType.XlTypeArray;
                pOper->arrayValue.Rows = rows;
                pOper->arrayValue.Columns = columns;
                pOper->arrayValue.pOpers = ((XlOper*)pNative + 1);
            }
            // This loop won't be entered in the empty-array case (rows * columns == 0)
            for (int i = 0; i < rows * columns; i++)
            {
                // Get the right object out of the array
                object obj;
                if (rank == 1)
                {
                    obj = ((object[])ManagedObj)[columnBase + i];
                }
                else
                {
                    int row = i / allColumns;
                    int column = i % allColumns;
                    obj = ((object[,])ManagedObj)[rowBase + row, columnBase + column];
                }

                // Get the right pOper
                pOper = (XlOper*)pNative + i + 1;

                // Set up the oper from the object
                if (obj is double)
                {
                    pOper->numValue = (double)obj;
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is string)
                {
                    // We count all of the string lengths, 
                    string str = (string)obj;
                    cbNativeStrings += Math.Min(str.Length, 255) + 1;
                    // mark the Oper as a string, and
                    // later allocate memory and return to fix pointers
                    pOper->xlType = XlType.XlTypeString;
                }
                else if (obj is DateTime)
                {
                    pOper->numValue = ((DateTime)obj).ToOADate();
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (IntegrationMarshalHelpers.IsExcelErrorObject(obj))
                {
                    pOper->errValue = (ushort)IntegrationMarshalHelpers.ExcelErrorGetValue(obj);
                    pOper->xlType = XlType.XlTypeError;
                }
                else if (IntegrationMarshalHelpers.IsExcelMissingObject(obj))
                {
                    pOper->xlType = XlType.XlTypeMissing;
                }
                else if (IntegrationMarshalHelpers.IsExcelEmptyObject(obj))
                {
                    pOper->xlType = XlType.XlTypeEmpty;
                }
                else if (obj is bool)
                {
                    pOper->boolValue = (bool)obj ? (ushort)1 : (ushort)0;
                    pOper->xlType = XlType.XlTypeBoolean;
                }
                else if (obj is short)
                {
                    pOper->numValue = (double)((short)obj); // int16 in XlOper
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is ushort)
                {
                    pOper->numValue = (double)((ushort)obj);
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is int)
                {
                    pOper->numValue = (double)((int)obj);
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is uint)
                {
                    pOper->numValue = (double)((uint)obj);
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is long)
                {
                    pOper->numValue = (double)((long)obj);
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is ulong)
                {
                    pOper->numValue = (double)((ulong)obj);
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is decimal)
                {
                    pOper->numValue = (double)((decimal)obj);
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (obj is float)
                {
                    pOper->numValue = (double)((float)obj);
                    pOper->xlType = XlType.XlTypeNumber;
                }
                else if (IntegrationMarshalHelpers.IsExcelReferenceObject(obj))
                {
                    pOper->xlType = XlType.XlTypeReference;
                    // First we count all of these, 
                    // later allocate memory and return to fix pointers
                    numReferenceOpers++;
                    numReferences += IntegrationMarshalHelpers.ExcelReferenceGetRectangleCount(obj); // ((ExcelReference)obj).InnerReferences.Count;
                }
                else if (obj is object[])
                {
                    XlObjectArrayMarshaler m = new XlObjectArrayMarshaler(1);
                    nestedInstances.Add(m);
                    XlOper* pNested = (XlOper*)m.MarshalManagedToNative(obj);
                    if (pNested->xlType == XlType.XlTypeArray)
                    {
                        pOper->xlType = XlType.XlTypeArray;
                        pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                        pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                        pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                    }
                    else
                    {
                        // This is the case where the array passed in has 0 length.
                        // We set to an error to at least have a valid XLOPER
                        pOper->xlType = XlType.XlTypeError;
                        pOper->errValue = (ushort)IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                    }
                }
                else if (obj is object[,])
                {
                    XlObjectArrayMarshaler m = new XlObjectArrayMarshaler(2);
                    nestedInstances.Add(m);
                    XlOper* pNested = (XlOper*)m.MarshalManagedToNative(obj);
                    if (pNested->xlType == XlType.XlTypeArray)
                    {
                        pOper->xlType = XlType.XlTypeArray;
                        pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                        pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                        pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                    }
                    else
                    {
                        // This is the case where the array passed in has 0 length.
                        // We set to an error to at least have a valid XLOPER
                        pOper->xlType = XlType.XlTypeError;
                        pOper->errValue = (ushort)IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                    }
                }
                else if (obj is System.Reflection.Missing)
                {
                    pOper->xlType = XlType.XlTypeMissing;
                }
                else if (obj == null)
                {
                    // DOCUMENT: I return Empty for nulls inside the Array, 
                    // which is not consistent with what happens in other settings.
                    // In particular not consistent with the results of the XlObjectMarshaler
                    // (which is not called when a null is returned,
                    // and interpreted as ExcelErrorNum in Excel)
                    // This works well for xlSet though.
                    // CONSIDER: Create an ExcelEmpty type to allow this to be more explicit,
                    // and return ErrNum here
                    pOper->xlType = XlType.XlTypeEmpty;
                }
                else
                {
                    // Default error return
                    pOper->errValue = (ushort)IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                    pOper->xlType = XlType.XlTypeError;
                }
            } // end of first pass

			// Now handle strings
			if (cbNativeStrings > 0)
			{
				// Allocate room for all the strings
                pNativeStrings = Marshal.AllocCoTaskMem(cbNativeStrings);
                // Go through the Opers and set each string
                byte* pCurrent = (byte*)pNativeStrings;
				for (int i = 0; i < rows * columns; i++)
				{
					// Get the corresponding oper
					pOper = (XlOper*)pNative + i + 1;
					if (pOper->xlType == XlType.XlTypeString)
					{
						// Get the string from the managed array
						string str;
						if (rank == 1)
						{
							str = (string)((object[])ManagedObj)[i];
						}
						else
						{
							int row = i / allColumns;
							int column = i % allColumns;
							str = (string)((object[,])ManagedObj)[rowBase + row, columnBase + column];
						}

                        XlString* pXlString = (XlString*)pCurrent;
						pOper->pstrValue = pXlString;
						int charCount = Math.Min(str.Length, 255);
						fixed (char* psrc = str)
						{
                            // Write the data and length to the XlString
                            // Support for system codepage by hmd
							// int written = Encoding.ASCII.GetBytes(psrc, charCount, pXlString->Data, 255);
							Encoding enc = Encoding.GetEncoding(ASCIIEncoding.Default.CodePage,
                                                                EncoderFallback.ReplacementFallback,
                                                                DecoderFallback.ReplacementFallback);
                            int written = enc.GetBytes(psrc, charCount, pXlString->Data, 255);
                            pXlString->Length = (byte)written;
                            // Increment pointer within allocated memory
                            pCurrent += written + 1;
                        }
					}
				}
			}

			// Now handle references
			if (numReferenceOpers > 0)
			{
				// Allocate room for all the references
                int cbNativeReferences = numReferenceOpers * sizeof(ushort) + numReferences * Marshal.SizeOf(typeof(XlOper.XlRectangle));
				pNativeReferences = Marshal.AllocCoTaskMem(cbNativeReferences);
				IntPtr pCurrent = pNativeReferences;
				// Go through the Opers and set each reference
				int refOperIndex = 0;
				for (int i = 0; i < rows * columns && refOperIndex < numReferenceOpers; i++)
				{
					// Get the corresponding oper
					pOper = (XlOper*)pNative + i + 1;
					if (pOper->xlType == XlType.XlTypeReference)
					{
						// Get the reference from the managed array
						object /*ExcelReference*/ r;
						if (rank == 1)
						{
							r = /*(ExcelReference)*/((object[])ManagedObj)[i];
						}
						else
						{
							int row = i / allColumns;
							int column = i % allColumns;
							r = /*(ExcelReference)*/((object[,])ManagedObj)[rowBase + row, columnBase + column];
						}

                        int refCount = IntegrationMarshalHelpers.ExcelReferenceGetRectangleCount(r); // r.InnerReferences.Count
						int numBytes = Marshal.SizeOf(typeof(ushort)) + refCount * Marshal.SizeOf(typeof(XlOper.XlRectangle));

						IntegrationMarshalHelpers.SetExcelReference(pOper, (XlOper.XlMultiRef*)pCurrent, r);

						pCurrent = new IntPtr(unchecked(pCurrent.ToInt32() + (int)numBytes));
						refOperIndex++;
					}
				}
			}
			
			if (!isExcel4v)
			{
				// For big allocations, ensure that Excel allows us to free the memory
                if (rows * columns * 16 + cbNativeStrings + numReferences * 8 > 65535)
					pOper->xlType |= XlType.XlBitDLLFree;

				// We are done
				return pNative;
			}
			else
			{
				// For the Excel4v call, we need to return an array
				// which will contain the pointers to the Opers.
                int cbOperPointers = columns * Marshal.SizeOf(typeof(XlOper*));
				pOperPointers = Marshal.AllocCoTaskMem(cbOperPointers);
				XlOper** pOpers = (XlOper**)pOperPointers;
				for (int i = 0; i < columns; i++)
				{
					pOpers[i] = (XlOper*)pNative + i + 1;
				}
				return pOperPointers;
			}
		}

		unsafe public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			ICustomMarshaler m = XlObjectMarshaler.GetInstance("");
			object managed = m.MarshalNativeToManaged(pNativeData);

			// Duplication here, because the types are different and wrapped in fixed blocks
			if (rank == 1)
			{
				if (managed == null || !(managed is object[,]))
				{
					return new object[1] { managed };
				}
				else // managed is object[,]: turn first row (or column) into object[]
				{
                    object[] array;
					object[,] all = (object[,])managed;
                    int rows = all.GetLength(0);
                    int columns = all.GetLength(1);

                    if (columns == 1)
                    {
                        // Take the one and only column as the array
                        array = new object[rows];
                        for (int i = 0; i < rows; i++)
                        {
                            array[i] = all[i, 0];
                        }
                    }
                    else
                    {
                        // Take first row only
                        array = new object[columns];
                        for (int j = 0; j < columns; j++)
                        {
                            array[j] = all[0, j];
                        }
                    }
					return array;
				}
			}
			else if (rank == 2)
			{
				if (managed == null || !(managed is object[,]))
				{
					return new object[,] { { managed } };
				}
				else // managed is object[,]
				{
					return managed;
				}
			}
			else
			{
				Debug.Fail("Damaged XlObjectArrayMarshaler rank");
				throw new InvalidOperationException("Damaged XlObjectArrayMarshaler rank");
			}
		}

		public void CleanUpManagedData(object ManagedObj) { }
		public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, never called as part of marshaling.
		public int GetNativeDataSize() { return -1; }
		
        // Implementation of IDisposable pattern
        // DOCUMENT: Not threadsafe implementation
        // Only called from Excel4v parameter marshaling
        private bool disposed = false;
       
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Also called to clean up the instance on every return call...
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                Reset(disposing);
            }
            disposed = true;
		}

        ~XlObjectArrayMarshaler()
        {
            Dispose(false);
        }

        // Called for disposal and for reset on every call to ManagedToNative.
        private void Reset(bool disposeNested)
        {
            if (disposeNested)
            {
                // Clean up the nested Instances
                foreach (XlObjectArrayMarshaler m in nestedInstances)
                {
                    m.Dispose();
                }
                nestedInstances.Clear();
            }

            Marshal.FreeCoTaskMem(pNative);
            pNative = IntPtr.Zero;

            Marshal.FreeCoTaskMem(pNativeStrings);
            pNativeStrings = IntPtr.Zero;

            Marshal.FreeCoTaskMem(pNativeReferences);
            pNativeReferences = IntPtr.Zero;

            Marshal.FreeCoTaskMem(pOperPointers);
            pOperPointers = IntPtr.Zero;
        }
	}

    // We would prefer to get a double, but need to take 
    // XlOper to ensure marshaling
    public unsafe class XlDecimalParameterMarshaler : ICustomMarshaler
    {
        static ICustomMarshaler instance;

        public XlDecimalParameterMarshaler()
        {
        }

        public static ICustomMarshaler GetInstance(string marshalCookie)
        {
            if (instance == null)
                instance = new XlDecimalParameterMarshaler();
            return instance;
        }

        public IntPtr MarshalManagedToNative(object ManagedObj)
        {
            // Not working in this direction at the moment
            throw new NotImplementedException("This marshaler only used for native to managed parameter marshaling.");
            //return null;
        }

        public object MarshalNativeToManaged(IntPtr pNativeData)
        {
            try
            {
                return (decimal)*((double*)pNativeData);
            }
            catch
            {
                // This case is where the range of the decimal is exceeded.
                // By returning null, the unboxing code will fail,
                // causing a runtime exception that is caught and returned as a #Value error.
                return null;
            }
        }

        public void CleanUpManagedData(object ManagedObj) { }
        public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
        public int GetNativeDataSize() { return -1; }
    }

    // We would prefer to get a double, but need to take 
    // XlOper to ensure marshaling
    public unsafe class XlLongParameterMarshaler : ICustomMarshaler
    {
        static ICustomMarshaler instance;

        public XlLongParameterMarshaler()
        {
        }

        public static ICustomMarshaler GetInstance(string marshalCookie)
        {
            if (instance == null)
                instance = new XlLongParameterMarshaler();
            return instance;
        }

        public IntPtr MarshalManagedToNative(object ManagedObj)
        {
            // Not working in this direction at the moment
            throw new NotImplementedException("This marshaler only used for native to managed parameter marshaling.");
            //return null;
        }

        public object MarshalNativeToManaged(IntPtr pNativeData)
        {
            try
            {
                return (long)*((double*)pNativeData);
            }
            catch
            {
                // This case is where the range of the long is exceeded.
                // By returning null, the unboxing code will fail,
                // causing a runtime exception that is caught and returned as a #Value error.
                return null;
            }
        }

        public void CleanUpManagedData(object ManagedObj) { }
        public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
        public int GetNativeDataSize() { return -1; }
    }

    internal unsafe static class IntegrationMarshalHelpers
    {
        static Type excelReferenceType;
        static Type excelErrorType;
        static ConstructorInfo excelReferenceConstructor;
        static MethodInfo excelReferenceAddReference;
        static PropertyInfo excelReferenceGetSheetId;
        static MethodInfo excelReferenceGetRectangleCount;
        static MethodInfo excelReferenceGetRectangles;

        static Type excelMissingType;
        static Type excelEmptyType;
        static object excelMissingValue;
        static object excelEmptyValue;

        internal static void Bind(Assembly integrationAssembly)
        {
            excelReferenceType = integrationAssembly.GetType("ExcelDna.Integration.ExcelReference");
            excelErrorType = integrationAssembly.GetType("ExcelDna.Integration.ExcelError");

            excelReferenceConstructor = excelReferenceType.GetConstructor( new Type[] { typeof(int), typeof(int), typeof(int), typeof(int), typeof(IntPtr) });
            excelReferenceAddReference = excelReferenceType.GetMethod("AddReference");
            excelReferenceGetSheetId = excelReferenceType.GetProperty("SheetId");
            excelReferenceGetRectangleCount = excelReferenceType.GetMethod("GetRectangleCount", BindingFlags.NonPublic | BindingFlags.Instance);
            excelReferenceGetRectangles = excelReferenceType.GetMethod("GetRectangles", BindingFlags.NonPublic | BindingFlags.Instance);

            excelMissingType = integrationAssembly.GetType("ExcelDna.Integration.ExcelMissing");
            excelEmptyType = integrationAssembly.GetType("ExcelDna.Integration.ExcelEmpty");

            FieldInfo excelMissingValueField = excelMissingType.GetField("Value", BindingFlags.Static | BindingFlags.Public);
            excelMissingValue = excelMissingValueField.GetValue(null);
            FieldInfo excelEmptyValueField = excelEmptyType.GetField("Value", BindingFlags.Static | BindingFlags.Public);
            excelEmptyValue = excelEmptyValueField.GetValue(null);
        }

        internal static object CreateExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast, IntPtr sheetId)
        {
            return excelReferenceConstructor.Invoke(new object[] { rowFirst, rowLast, columnFirst, columnLast, sheetId });
        }

        internal unsafe static void SetExcelReference(XlOper* pOper, XlOper.XlMultiRef* pMultiRef, object /*ExcelReference*/ r)
        {
            IntPtr sheetId = ExcelReferenceGetSheetId(r);
            int[][] rects = ExcelReferenceGetRectangles(r);
            int rectCount = rects.GetLength(0);

            pOper->xlType = XlType.XlTypeReference;
            pOper->refValue.SheetId = sheetId;

            pOper->refValue.pMultiRef = pMultiRef;
            pOper->refValue.pMultiRef->Count = (ushort)rectCount;

            XlOper.XlRectangle* pRectangles = (XlOper.XlRectangle*)(&pOper->refValue.pMultiRef->Rectangles);

            for (int i = 0; i < rectCount; i++)
            {
                pRectangles[i].RowFirst = (ushort)rects[i][0];
                pRectangles[i].RowLast = (ushort)rects[i][1];
                pRectangles[i].ColumnFirst = (byte)rects[i][2];
                pRectangles[i].ColumnLast = (byte)rects[i][3];
            }
        }

        internal unsafe static void SetExcelReference12(XlOper12* pOper, XlOper12.XlMultiRef12* pMultiRef, object /*ExcelReference*/ r)
        {
            IntPtr sheetId = ExcelReferenceGetSheetId(r);
            int[][] rects = ExcelReferenceGetRectangles(r);
            int rectCount = rects.GetLength(0);

            pOper->xlType = XlType12.XlTypeReference;
            pOper->refValue.SheetId = sheetId;

            pOper->refValue.pMultiRef = pMultiRef;
            pOper->refValue.pMultiRef->Count = (ushort)rectCount;

            XlOper12.XlRectangle12* pRectangles = (XlOper12.XlRectangle12*)(&pOper->refValue.pMultiRef->Rectangles);

            for (int i = 0; i < rectCount; i++)
            {
                pRectangles[i].RowFirst = rects[i][0];
                pRectangles[i].RowLast = rects[i][1];
                pRectangles[i].ColumnFirst = rects[i][2];
                pRectangles[i].ColumnLast = rects[i][3];
            }
        }

        internal static void ExcelReferenceAddReference(object r, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            excelReferenceAddReference.Invoke(r, new object[] { rowFirst, rowLast, columnFirst, columnLast });
        }

        internal static IntPtr ExcelReferenceGetSheetId(object r)
        {
            return (IntPtr)excelReferenceGetSheetId.GetValue(r, null);
        }

        internal static int ExcelReferenceGetRectangleCount(object r)
        {
            return (int)excelReferenceGetRectangleCount.Invoke(r, null);
        }

        internal static int[][] ExcelReferenceGetRectangles(object r)
        {
            return (int[][])excelReferenceGetRectangles.Invoke(r, null);
        }

        internal static bool IsExcelReferenceObject(object o)
        {
            return excelReferenceType.IsInstanceOfType(o);
        }

        internal static bool IsExcelErrorObject(object o)
        {
            return excelErrorType.IsInstanceOfType(o);
        }

        internal static bool IsExcelMissingObject(object o)
        {
            return excelMissingType.IsInstanceOfType(o);
        }

        internal static bool IsExcelEmptyObject(object o)
        {
            return excelEmptyType.IsInstanceOfType(o);
        }

        internal static int ExcelErrorGetValue(object e)
        {
            return (int)(ushort)e;
        }

        internal static Type GetExcelErrorType()
        {
            return excelErrorType;
        }

        internal static object GetExcelErrorObject(int errorCode)
        {
            return Enum.ToObject(excelErrorType, errorCode);
        }

        internal const int ExcelError_ExcelErrorValue = 15;

        internal static object GetExcelMissingValue()
        {
            return excelMissingValue;
        }

        internal static object GetExcelEmptyValue()
        {
            return excelEmptyValue;
        }
    }
}
