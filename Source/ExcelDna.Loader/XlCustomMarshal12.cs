/*
  Copyright (C) 2005-2009 Govert van Drimmelen

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

// WARNING: Memory allocation needs fixing before any attempt at MultiThreading.
//          Probably allocate instances per thread.... ?

// TODO: Check what happens for re-entrancy, e.g. Calling a UDF from Excel.Excel4 !!

// TODO: Marshalers should implement disposable pattern.

namespace ExcelDna.Loader
{
	// Internal Implementations of the Excel Types
	// CONSIDER: How to (if?) make these available to the user code
	// For now I think of this as an internal structure used in the marshaling

    [StructLayout(LayoutKind.Explicit)]
    internal unsafe struct XlString12
    {
        [FieldOffset(0)]
        public ushort Length;
        [FieldOffset(2)]
        public fixed char Data[1]; // Actually Data[Length]

        public static readonly int MaxLength = 32767; // chars
    }
    [Flags]
    internal enum XlType12 : uint
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
	internal unsafe struct XlOper12
	{
        [StructLayout(LayoutKind.Explicit)]
		unsafe public struct XlOper12Array
		{
			[FieldOffset(0)]
			public XlOper12* pOpers;
			[FieldOffset(4)]
			public int Rows;
			[FieldOffset(8)]
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
			public XlRectangle12 Rectangles;	// Not XlRectangle12, actually Rectangles[Count] !
		}

		[StructLayout(LayoutKind.Explicit)]
		unsafe public struct XlSReference12
		{
			[FieldOffset(0)]
			public ushort Count;                // Always = 1
			[FieldOffset(4)]
			public XlRectangle12 Reference;
		}

		[StructLayout(LayoutKind.Explicit)]
		unsafe public struct XlReference12
		{
			[FieldOffset(0)]
			public XlMultiRef12* pMultiRef;
			[FieldOffset(4)]
			public uint SheetId;
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
		[FieldOffset(24)]
		public XlType12 xlType;
    }

	public class XlString12ReturnMarshaler : ICustomMarshaler
	{
        [ThreadStatic] static ICustomMarshaler instance;
		IntPtr pNative; // Pointer to XlString, allocated once on initialization

		public XlString12ReturnMarshaler()
		{
            int size = Marshal.SizeOf(typeof(XlString12)) + ((XlString12.MaxLength - 1) /* 1 char is in Data[1] */ * 2 /* 2 bytes per char */);
			pNative = Marshal.AllocCoTaskMem(size);
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			if (instance == null)
				instance = new XlString12ReturnMarshaler();
			return instance;
		}

		unsafe public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			// CONSIDER: Checking for null, checking object type.
			// CONSIDER: Marshal back as OPER for errors etc.

			// DOCUMENT: This function is not called if the return is null!
			// DOCUMENT: A null pointer is immediately returned to Excel, resulting in #NUM!

			String src = (String)ManagedObj;
			XlString12* pdest = (XlString12*)pNative;
            ushort charCount = (ushort)Math.Min(src.Length, XlString12.MaxLength);
            fixed (char* psrc = src)
            {
                char* ps = psrc;
                char* pd = pdest->Data;
                for (int k = 0; k < charCount; k++)
                {
                    *(pd++) = *(ps++);
                }
            }
            pdest->Length = charCount;
			
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
	public unsafe class XlBoolean12ReturnMarshaler : ICustomMarshaler
	{
        [ThreadStatic] static ICustomMarshaler instance;
		IntPtr pNative; // this is really an XlOper, and is is allocated once, 
						// when the marshaller is constructed, 
						// and is never reclaimed

		public XlBoolean12ReturnMarshaler()
		{
			int size = Marshal.SizeOf(typeof(XlOper12));
			pNative = Marshal.AllocCoTaskMem(size);
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			if (instance == null)
				instance = new XlBoolean12ReturnMarshaler();
			return instance;

		}

		public IntPtr MarshalManagedToNative(object ManagedObj)
		{
			XlOper12* xlOper = (XlOper12*)pNative;
			xlOper->boolValue = (bool)ManagedObj ? 1 : 0;
			xlOper->xlType = XlType12.XlTypeBoolean;
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
	internal unsafe struct XlFP12
	{
		public int Rows;
		public int Columns;
        public fixed double Values[1]; // Actually, Values[rows][columns]
	}

    public class XlDoubleArray12Marshaler : ICustomMarshaler
	{

		// CONSIDER: Marshal all return types of double[,] as XLOPER
		// and set xlFree bit, and handle callback.
		// This will reduce memory usage but be slower, as we would get callback
		// into managed code, unless we implement xlFree in native
		// (we can use Com memory allocator to free there)
		// For now just do fast, simple, slightly memory hogging thing.

        [ThreadStatic] static XlDoubleArray12Marshaler instance1;	// For rank 1 arrays
        [ThreadStatic] static XlDoubleArray12Marshaler instance2;	// For rank 2 arrays

		int rank;
		IntPtr pNative; // For managed -> native returns 
						// This points to the last FP that was marshaled.
						// FPs are re-allocated on every managed->native transition

		public XlDoubleArray12Marshaler(int rank)
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
					instance1 = new XlDoubleArray12Marshaler(1);
				return instance1;
			}
			else if (marshalCookie == "2")
			{
				if (instance2 == null)
					instance2 = new XlDoubleArray12Marshaler(2);
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
			
			int rows;
			int columns;
			if (rank == 1)
			{
				double[] doubles = (double[])ManagedObj;

				rows = 1;
				columns = doubles.Length;

				fixed(double* src = doubles)
				{
					AllocateFP12AndCopy(src, rows, columns);	
				}
			}
			else if (rank == 2)
			{
				double[,] doubles = (double[,])ManagedObj;

				rows = doubles.GetLength(0);
                columns = doubles.GetLength(1);

				fixed (double* src = doubles)
				{
					AllocateFP12AndCopy(src, rows, columns);
				}
			}
			else
			{
				throw new InvalidOperationException("Damaged XlDoubleArrayMarshaler rank");
			}

            // CONSIDER: If large, mark and deal with xlDllFree

			return pNative;
		}

		unsafe private void AllocateFP12AndCopy(double* pSrc, int rows, int columns)
		{
            XlFP12* pFP;
            if (columns == 0)
            {
                // TODO: Review handling of this corner case
                pNative = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(XlFP12)));
                pFP = (XlFP12*)pNative;
                pFP->Rows = 1;
                pFP->Columns = 1;
                pFP->Values[0] = 0;
                return;
            }
			int size = Marshal.SizeOf(typeof(XlFP12)) +
			   Marshal.SizeOf(typeof(double)) * (rows * columns - 1); // room for one double is already in FP12 struct
			pNative = Marshal.AllocCoTaskMem(size);

			pFP = (XlFP12*)pNative;
			pFP->Rows = rows;
			pFP->Columns = columns;
			int count = rows * columns;
			// Fast copy
			CopyDoubles(pSrc, pFP->Values, count);				
		}

		unsafe public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			object result;
			XlFP12* pFP = (XlFP12*)pNativeData;
			
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
				Debug.Fail("Damaged XlDoubleArray12Marshaler rank");
				throw new InvalidOperationException("Damaged XlDoubleArray12Marshaler rank");
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

    public class XlObject12Marshaler : ICustomMarshaler
	{
        // Shared instance used for all return values
		[ThreadStatic] static ICustomMarshaler instance;

		IntPtr pNative; // this is really an XlOper, and it is allocated once, 
                        // when the marshaller is constructed.

		public XlObject12Marshaler()
		{
			int size = Marshal.SizeOf(typeof(XlOper12));
			pNative = Marshal.AllocCoTaskMem(size);
		}

		public static ICustomMarshaler GetInstance(string marshalCookie)
		{
			if (instance == null)
				instance = new XlObject12Marshaler();
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
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->numValue = (double)ManagedObj;
                pOper->xlType = XlType12.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is string)
			{
				// TODO: Consolidate these?
				ICustomMarshaler m = XlString12ReturnMarshaler.GetInstance("");
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->pstrValue = (XlString12*)m.MarshalManagedToNative(ManagedObj);
                pOper->xlType = XlType12.XlTypeString;
				return pNative;
			}
			else if (ManagedObj is DateTime)
			{
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->numValue = ((DateTime)ManagedObj).ToOADate();
                pOper->xlType = XlType12.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is bool)
			{
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->boolValue = (bool)ManagedObj ? (ushort)1 : (ushort)0;
                pOper->xlType = XlType12.XlTypeBoolean;
				return pNative;
			}
			else if (ManagedObj is object[])
			{
				// Redirect to the ObjectArray Marshaler
				// CONSIDER: This might cause some memory to get stuck, 
				// since the memory for the array marshaler is not the same as for this
				ICustomMarshaler m = XlObjectArray12Marshaler.GetInstance("1");
				return m.MarshalManagedToNative(ManagedObj);
			}
			else if (ManagedObj is object[,])
			{
				// Redirect to the ObjectArray Marshaler
				// CONSIDER: This might cause some memory to get stuck, 
				// since the memory for the array marshaler is not the same as for this
                ICustomMarshaler m = XlObjectArray12Marshaler.GetInstance("2");
                IntPtr native = m.MarshalManagedToNative(ManagedObj);
                return native;
			}
			else if (ManagedObj is double[])
			{
				double[] doubles = (double[])ManagedObj;
				object[] objects = new object[doubles.Length];
				Array.Copy(doubles, objects, doubles.Length);
				ICustomMarshaler m = XlObjectArray12Marshaler.GetInstance("1");
				return m.MarshalManagedToNative(objects);
			}
			else if (ManagedObj is double[,])
			{
				double[,] doubles = (double[,])ManagedObj;
				object[,] objects = new object[doubles.GetLength(0), doubles.GetLength(1)];
				Array.Copy(doubles, objects, doubles.GetLength(0) * doubles.GetLength(1));
				ICustomMarshaler m = XlObjectArray12Marshaler.GetInstance("2");
				return m.MarshalManagedToNative(objects);
			}
			else if (IntegrationMarshalHelpers.IsExcelErrorObject(ManagedObj))
			{
				XlOper12* pOper = (XlOper12*)pNative;
                pOper->errValue = IntegrationMarshalHelpers.ExcelErrorGetValue(ManagedObj);
                pOper->xlType = XlType12.XlTypeError;
				return pNative;
			}
            else if (IntegrationMarshalHelpers.IsExcelMissingObject(ManagedObj))
            {
                XlOper12* pOper = (XlOper12*)pNative;
                pOper->xlType = XlType12.XlTypeMissing;
                return pNative;
            }
            else if (IntegrationMarshalHelpers.IsExcelEmptyObject(ManagedObj))
            {
                XlOper12* pOper = (XlOper12*)pNative;
                pOper->xlType = XlType12.XlTypeEmpty;
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
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->numValue = (double)((short)ManagedObj);
				pOper->xlType = XlType12.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is System.Reflection.Missing)
			{
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->xlType = XlType12.XlTypeMissing;
				return pNative;
			}
            else if (ManagedObj is int)
			{
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->numValue = (double)((int)ManagedObj);
				pOper->xlType = XlType12.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is uint)
			{
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->numValue = (double)((uint)ManagedObj);
				pOper->xlType = XlType12.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is byte)
			{
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->numValue = (double)((byte)ManagedObj);
				pOper->xlType = XlType12.XlTypeNumber;
				return pNative;
			}
			else if (ManagedObj is ushort)
			{
				XlOper12* pOper = (XlOper12*)pNative;
				pOper->numValue = (double)((ushort)ManagedObj);
				pOper->xlType = XlType12.XlTypeNumber;
				return pNative;
			}
            else if (ManagedObj is decimal)
            {
                XlOper12* pOper = (XlOper12*)pNative;
                pOper->numValue = (double)((decimal)ManagedObj);
                pOper->xlType = XlType12.XlTypeNumber;
                return pNative;
            }
            // 13 November 2006 -- this marshaler is not used to set up the Excel4 return value anymore.
            //else if (ManagedObj == null)
            //{
            //    // This is never the case for regular marshaling, only for 
            //    // return value for Excel4 function
            //    XlOper* pOper = (XlOper*)pNative;
            //    pOper->xlType = XlType12.XlTypeEmpty;
            //    return pNative;
            //}
            else
            {
                // Default error return
                XlOper12* pOper = (XlOper12*)pNative;
                pOper->errValue = IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                pOper->xlType = XlType12.XlTypeError;
                return pNative;
            }
		}

		unsafe public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			// Make a nice object from the native OPER
			object managed;
			XlOper12* pOper = (XlOper12*)pNativeData;
			// Ignore any Free flags
            XlType12 type = pOper->xlType & ~XlType12.XlBitXLFree & ~XlType12.XlBitDLLFree;
			switch (type)
			{
				case XlType12.XlTypeNumber:
					managed = pOper->numValue;
					break;
				case XlType12.XlTypeString:
                    XlString12* pString = pOper->pstrValue;
                    managed = new string(pString->Data, 0, pString->Length);
					break;
				case XlType12.XlTypeBoolean:
					managed = pOper->boolValue == 1;
					break;
				case XlType12.XlTypeError:
					managed = IntegrationMarshalHelpers.GetExcelErrorObject(pOper->errValue);
					break;
                case XlType12.XlTypeMissing:
                    // DOCUMENT: Changed in version 0.17.
                    // managed = System.Reflection.Missing.Value;
                    managed = IntegrationMarshalHelpers.GetExcelMissingValue();
                    break;
                case XlType12.XlTypeEmpty:
                    // DOCUMENT: Changed in version 0.17.
                    // managed = null;
                    managed = IntegrationMarshalHelpers.GetExcelEmptyValue();
                    break;
                case XlType12.XlTypeArray:
					int rows = pOper->arrayValue.Rows;
					int columns = pOper->arrayValue.Columns;
					object[,] array = new object[rows, columns];
					XlOper12* opers = (XlOper12*)pOper->arrayValue.pOpers;
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
				case XlType12.XlTypeReference:
					object /*ExcelReference*/ r;
					if (pOper->refValue.pMultiRef == (XlOper12.XlMultiRef12*)IntPtr.Zero)
					{
						r = IntegrationMarshalHelpers.CreateExcelReference(0, 0, 0, 0, pOper->refValue.SheetId);
					}
					else
					{
						ushort numAreas = *(ushort*)pOper->refValue.pMultiRef;
						XlOper12.XlRectangle12* pAreas = (XlOper12.XlRectangle12*)((uint)pOper->refValue.pMultiRef + 4 /* FieldOffset for XlRectangles */);
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
				case XlType12.XlTypeSReference:
                    uint sheetId = XlCallImpl.GetCurrentSheetId12();
					object /*ExcelReference*/ sref;
                    sref = IntegrationMarshalHelpers.CreateExcelReference(
                                            pOper->srefValue.Reference.RowFirst,
											pOper->srefValue.Reference.RowLast,
											pOper->srefValue.Reference.ColumnFirst,
											pOper->srefValue.Reference.ColumnLast, 
											sheetId /*Current sheet (not active sheet)*/);
                    managed = sref;
					break;
                case XlType12.XlTypeInt: // Never passed from Excel to a UDF! int32 in XlOper12
                    managed = (double)pOper->intValue;
                    break;
                default:
                    // unheard of !!
					managed = null;
					break;
			}
			return managed;
		}

		public void CleanUpManagedData(object ManagedObj) { }
        public void CleanUpNativeData(IntPtr pNativeData) { } // Can't do anything useful here, as the managed to native marshaling is for a return parameter.
		public int GetNativeDataSize() { return -1; }
	}

    public class XlObjectArray12Marshaler : ICustomMarshaler, IDisposable
	{
		// CONSIDER: Marshal return types of object[,] as XLOPER12
		// and set xlFree bit, and handle callback.
		// This will reduce memory usage but be slower, as we would get callback
		// into managed code, unless we implement xlFree in native
		// (we can use Com memory allocator to free there)
		// For now just do fast, simple, slightly memory hogging thing.

        [ThreadStatic] static XlObjectArray12Marshaler instance1;	// For rank 1 arrays
        [ThreadStatic] static XlObjectArray12Marshaler instance2;	// For rank 2 arrays

        int rank;
        List<XlObjectArray12Marshaler> nestedInstances = new List<XlObjectArray12Marshaler>();
		bool isExcel12v;	// Used for calls to Excel12 -- flags that returned native data should look different

		IntPtr pNative; // For managed -> native returns 
		// This points to the last OPER (and contained OPER array) that was marshaled
		// OPERs are re-allocated on every managed->native transition
		IntPtr pNativeStrings;
		IntPtr pNativeReferences;
		
        IntPtr pOperPointers; // Used for calls to Excel4v - points to the array of oper addresses

        public XlObjectArray12Marshaler()
        {
            this.rank = 0;  // Must be set before use.
            this.isExcel12v = false;
        }

		public XlObjectArray12Marshaler(int rank)
		{
            this.rank = rank;
            this.isExcel12v = false;
		}

        public XlObjectArray12Marshaler(int rank, bool isExcel12v)
        {
            this.rank = rank;
            this.isExcel12v = isExcel12v;
        }

		public static void FreeMemory()
		{
            // This method is only called via AutoFree for an instance 
            // instanceX are ThreadStatic
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
                    instance1 = new XlObjectArray12Marshaler(1);
                return instance1;
            }
            else if (marshalCookie == "2")
            {
                if (instance2 == null)
                    instance2 = new XlObjectArray12Marshaler(2);
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

			// TODO: Remove duplication - due to fixed / pointer interaction
            // TODO: Might manages strings differently - currently I allocate the maximum length of 255 bytes
            //          for each string. Instead, I might just allocate the required number of bytes.

			Reset(true);

			int rows;
			int columns; // those in the returned array
			if (rank == 1)
			{
				object[] objects = (object[])ManagedObj;

				rows = 1;
				columns = objects.Length;
			}
			else if (rank == 2)
			{
				object[,] objects = (object[,])ManagedObj;

				rows = objects.GetLength(0);
				columns = objects.GetLength(1);
			}
			else
			{
				throw new InvalidOperationException("Damaged XlObjectArrayMarshaler rank");
			}

            int cbNativeStrings = 0;
			int numReferenceOpers = 0;
			int numReferences = 0;

			// Allocate native space
			int cbNative =  Marshal.SizeOf(typeof(XlOper12)) +				// OPER that is returned
				            Marshal.SizeOf(typeof(XlOper12)) * (rows * columns);	// Array of OPER inside the result
			pNative = Marshal.AllocCoTaskMem(cbNative);

			// Set up returned OPER
			XlOper12* pOper = (XlOper12*)pNative;
			pOper->xlType = XlType12.XlTypeArray;
			pOper->arrayValue.Rows = rows;
			pOper->arrayValue.Columns = columns;
			pOper->arrayValue.pOpers = ((XlOper12*)pNative + 1);

			for (int i = 0; i < rows * columns; i++)
			{
				// Get the right object out of the array
				object obj;
				if (rank == 1)
				{
					obj = ((object[])ManagedObj)[i];
				}
				else
				{
					int row = i / columns;
					int column = i % columns;
					obj = ((object[,])ManagedObj)[row, column];
				}

				// Get the right pOper
				pOper = (XlOper12*)pNative + i + 1;

				// Set up the oper from the object
				if (obj is double)
				{
					pOper->numValue = (double)obj;
                    pOper->xlType = XlType12.XlTypeNumber;
				}
				else if (obj is string)
				{
					// We count all of the string lengths, 
                    string str = (string)obj;
                    cbNativeStrings += (Marshal.SizeOf(typeof(XlString12)) + ((Math.Min(str.Length, XlString12.MaxLength) - 1) /* 1 char already in XlString */) * 2 /* 2 bytes per char */);
                    // mark the Oper as a string, and
                    // later allocate memory and return to fix pointers
                    pOper->xlType = XlType12.XlTypeString;
                }
				else if (obj is DateTime)
				{
					pOper->numValue = ((DateTime)obj).ToOADate();
                    pOper->xlType = XlType12.XlTypeNumber;
				}
				else if (IntegrationMarshalHelpers.IsExcelErrorObject(obj))
				{
					pOper->errValue = IntegrationMarshalHelpers.ExcelErrorGetValue(obj);
                    pOper->xlType = XlType12.XlTypeError;
				}
                else if (IntegrationMarshalHelpers.IsExcelMissingObject(obj))
                {
                    pOper->xlType = XlType12.XlTypeMissing;
                }
                else if (IntegrationMarshalHelpers.IsExcelEmptyObject(obj))
                {
                    pOper->xlType = XlType12.XlTypeEmpty;
                }
				else if (obj is bool)
				{
					pOper->boolValue = (bool)obj ? 1 : 0;
                    pOper->xlType = XlType12.XlTypeBoolean;
				}
                else if (obj is byte)
                {
                    pOper->numValue = (double)((byte)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is sbyte)
                {
                    pOper->numValue = (double)((sbyte)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is short)
				{
					pOper->numValue = (double)((short)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
				}
				else if (obj is ushort)
				{
					pOper->numValue = (double)((ushort)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
				}
				else if (obj is int)
				{
					pOper->numValue = (double)((int)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
				}
                else if (obj is uint)
                {
                    pOper->numValue = (double)((uint)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is long)
                {
                    pOper->numValue = (double)((long)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is ulong)
                {
                    pOper->numValue = (double)((long)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is decimal)
                {
                    pOper->numValue = (double)((decimal)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (obj is float)
                {
                    pOper->numValue = (double)((float)obj);
                    pOper->xlType = XlType12.XlTypeNumber;
                }
                else if (IntegrationMarshalHelpers.IsExcelReferenceObject(obj))
                {
                    pOper->xlType = XlType12.XlTypeReference;
                    // First we count all of these, 
                    // later allocate memory and return to fix pointers
                    numReferenceOpers++;
                    numReferences += IntegrationMarshalHelpers.ExcelReferenceGetRectangleCount(obj); // ((ExcelReference)obj).InnerReferences.Count;
                }
                else if (obj is object[])
                {
                    XlObjectArray12Marshaler m = new XlObjectArray12Marshaler(1);
                    nestedInstances.Add(m);
                    XlOper12* pNested = (XlOper12*)m.MarshalManagedToNative(obj);
                    pOper->xlType = XlType12.XlTypeArray;
                    pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                    pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                    pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                }
                else if (obj is object[,])
                {
                    XlObjectArray12Marshaler m = new XlObjectArray12Marshaler(2);
                    nestedInstances.Add(m);
                    XlOper12* pNested = (XlOper12*)m.MarshalManagedToNative(obj);
                    pOper->xlType = XlType12.XlTypeArray;
                    pOper->arrayValue.Rows = pNested->arrayValue.Rows;
                    pOper->arrayValue.Columns = pNested->arrayValue.Columns;
                    pOper->arrayValue.pOpers = pNested->arrayValue.pOpers;
                }
                else if (obj is Missing)
                {
                    pOper->xlType = XlType12.XlTypeMissing;
                }
                else if (obj == null)
                {
                    // DOCUMENT: I return Empty for nulls inside the Array, 
                    // which is not consistent with what happens in other settings.
                    // In particular not consistent with the results of the XlObjectMarshaler
                    // (which is not called when a null is returned,
                    // and interpreted as ExcelErrorNum in Excel)
                    // This works well for xlSet though.
                    pOper->xlType = XlType12.XlTypeEmpty;
                }
                else
                {
                    // Default error return
                    pOper->errValue = IntegrationMarshalHelpers.ExcelError_ExcelErrorValue;
                    pOper->xlType = XlType12.XlTypeError;
                }
			} // end of first pass

			// Now handle strings
			if (cbNativeStrings > 0)
			{
				// Allocate room for all the strings
                pNativeStrings = Marshal.AllocCoTaskMem(cbNativeStrings);
                // Go through the Opers and set each string
                char* pCurrent = (char*)pNativeStrings;
				for (int i = 0; i < rows * columns; i++)
				{
					// Get the corresponding oper
					pOper = (XlOper12*)pNative + i + 1;
                    if (pOper->xlType == XlType12.XlTypeString)
					{
						// Get the string from the managed array
						string str;
						if (rank == 1)
						{
							str = (string)((object[])ManagedObj)[i];
						}
						else
						{
							int row = i / columns;
							int column = i % columns;
							str = (string)((object[,])ManagedObj)[row, column];
						}

                        XlString12* pdest = (XlString12*)pCurrent;
						pOper->pstrValue = pdest;
						ushort charCount = (ushort)Math.Min(str.Length, XlString12.MaxLength);
						fixed (char* psrc = str)
						{
                            char* ps = psrc;
                            char* pd = pdest->Data;
                            for (int k = 0; k < charCount; k++)
                            {
                                *(pd++) = *(ps++);
                            }
                        }
                        pdest->Length = charCount;
                        // Increment pointer within allocated memory
                        pCurrent += charCount + 1;
					}
				}
			}

			// Now handle references
			if (numReferenceOpers > 0)
			{
				// Allocate room for all the references
                int cbNativeReferences = numReferenceOpers * 4 /* sizeof ushort + packing to get to field offset */ 
                                         + numReferences * Marshal.SizeOf(typeof(XlOper12.XlRectangle12));
				pNativeReferences = Marshal.AllocCoTaskMem(cbNativeReferences);
				IntPtr pCurrent = pNativeReferences;
				// Go through the Opers and set each reference
				int refOperIndex = 0;
				for (int i = 0; i < rows * columns && refOperIndex < numReferenceOpers; i++)
				{
					// Get the corresponding oper
					pOper = (XlOper12*)pNative + i + 1;
                    if (pOper->xlType == XlType12.XlTypeReference)
					{
						// Get the reference from the managed array
						object /*ExcelReference*/ r;
						if (rank == 1)
						{
							r = /*(ExcelReference)*/((object[])ManagedObj)[i];
						}
						else
						{
							int row = i / columns;
							int column = i % columns;
							r = /*(ExcelReference)*/((object[,])ManagedObj)[row, column];
						}

                        int refCount = IntegrationMarshalHelpers.ExcelReferenceGetRectangleCount(r); // r.InnerReferences.Count
						int numBytes = 4 /* sizeof ushort + packing to get to field offset */  
                                       + refCount * Marshal.SizeOf(typeof(XlOper12.XlRectangle12));

						IntegrationMarshalHelpers.SetExcelReference12(pOper, (XlOper12.XlMultiRef12*)pCurrent, r);

						pCurrent = (IntPtr)((uint)pCurrent + numBytes);
						refOperIndex++;
					}
				}
			}
			
			if (!isExcel12v)
			{
				// For big allocations, ensure that Excel allows us to free the memory
                if (rows * columns * 16 + cbNativeStrings + numReferences * 16 > 65535)
                    pOper->xlType |= XlType12.XlBitDLLFree;

				// We are done
				return pNative;
			}
			else
			{
				// For the Excel12v call, we need to return an array
				// which will contain the pointers to the Opers.
                int cbOperPointers = columns * Marshal.SizeOf(typeof(XlOper12*));
				pOperPointers = Marshal.AllocCoTaskMem(cbOperPointers);
				XlOper12** pOpers = (XlOper12**)pOperPointers;
				for (int i = 0; i < columns; i++)
				{
					pOpers[i] = (XlOper12*)pNative + i + 1;
				}
				return pOperPointers;
			}
		}

		unsafe public object MarshalNativeToManaged(IntPtr pNativeData)
		{
			ICustomMarshaler m = XlObject12Marshaler.GetInstance("");
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
				Debug.Fail("Damaged XlObjectArray12Marshaler rank");
				throw new InvalidOperationException("Damaged XlObjectArray12Marshaler rank");
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

        ~XlObjectArray12Marshaler()
        {
            Dispose(false);
        }

        // Called for disposal and for reset on every call to ManagedToNative.
        private void Reset(bool disposeNested)
        {
            if (disposeNested)
            {
                // Clean up the nested Instances
                foreach (XlObjectArray12Marshaler m in nestedInstances)
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
    public unsafe class XlDecimalParameter12Marshaler : ICustomMarshaler
    {
        [ThreadStatic] static ICustomMarshaler instance;

        public XlDecimalParameter12Marshaler()
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
}
