//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Diagnostics;
using System.Reflection;

namespace ExcelDna.Loader
{
    // There are some types in ExcelDna.Integration that we need to deal with to do marshaling.
    // Since we are not taking a reference to ExcelDna.Integration (for caution when doing unmanaged loading)
    // we access those types via reflection.
    internal unsafe static class IntegrationMarshalHelpers
    {
        static Type excelReferenceType;
        static ConstructorInfo excelReferenceConstructor;
        static ConstructorInfo excelReferenceConstructorRects;
        static PropertyInfo excelReferenceGetSheetId;
        static MethodInfo excelReferenceGetRectangleCount;
        static MethodInfo excelReferenceGetRectangles;

        static Type excelErrorType;

        static Type excelMissingType;
        static object excelMissingValue;

        static Type excelEmptyType;
        static object excelEmptyValue;

        static Type excelAsyncHandleType;
        static Type excelAsyncHandleNativeType;
        static ConstructorInfo excelAsyncHandleNativeConstructor;
        static FieldInfo excelAsyncHandleNativeHandleField;

        internal static void Bind(Assembly integrationAssembly)
        {
            excelReferenceType = integrationAssembly.GetType("ExcelDna.Integration.ExcelReference");
            excelReferenceConstructor = excelReferenceType.GetConstructor(new Type[] { typeof(int), typeof(int), typeof(int), typeof(int), typeof(IntPtr) });
            excelReferenceConstructorRects = excelReferenceType.GetConstructor(new Type[] { typeof(int[][]), typeof(IntPtr) });
            excelReferenceGetSheetId = excelReferenceType.GetProperty("SheetId");
            excelReferenceGetRectangleCount = excelReferenceType.GetMethod("GetRectangleCount", BindingFlags.NonPublic | BindingFlags.Instance);
            excelReferenceGetRectangles = excelReferenceType.GetMethod("GetRectangles", BindingFlags.NonPublic | BindingFlags.Instance);

            excelMissingType = integrationAssembly.GetType("ExcelDna.Integration.ExcelMissing");
            FieldInfo excelMissingValueField = excelMissingType.GetField("Value", BindingFlags.Static | BindingFlags.Public);
            excelMissingValue = excelMissingValueField.GetValue(null);

            excelEmptyType = integrationAssembly.GetType("ExcelDna.Integration.ExcelEmpty");
            FieldInfo excelEmptyValueField = excelEmptyType.GetField("Value", BindingFlags.Static | BindingFlags.Public);
            excelEmptyValue = excelEmptyValueField.GetValue(null);

            excelErrorType = integrationAssembly.GetType("ExcelDna.Integration.ExcelError");

            excelAsyncHandleType = integrationAssembly.GetType("ExcelDna.Integration.ExcelAsyncHandle");
            excelAsyncHandleNativeType = integrationAssembly.GetType("ExcelDna.Integration.ExcelAsyncHandleNative");
            excelAsyncHandleNativeConstructor = excelAsyncHandleNativeType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new Type[] { typeof(IntPtr) }, null);
            excelAsyncHandleNativeHandleField = excelAsyncHandleNativeType.GetField("_handle", BindingFlags.Instance | BindingFlags.NonPublic);

            Debug.Assert( excelReferenceType != null &&
                          excelReferenceConstructor != null &&
                          excelReferenceConstructorRects != null &&
                          excelReferenceGetSheetId != null &&
                          excelReferenceGetRectangleCount != null &&
                          excelReferenceGetRectangles != null &&

                          excelErrorType != null &&

                          excelMissingType != null &&
                          excelMissingValue != null &&

                          excelEmptyType != null &&
                          excelEmptyValue != null &&

                          excelAsyncHandleNativeType != null &&
                          excelAsyncHandleNativeConstructor != null &&
                          excelAsyncHandleNativeHandleField != null);
        }

        #region ExcelReference
        internal static bool IsExcelReferenceObject(object o)
        {
            return excelReferenceType.IsInstanceOfType(o);
        }

        internal static object CreateExcelReference(int rowFirst, int rowLast, int columnFirst, int columnLast, IntPtr sheetId)
        {
            return excelReferenceConstructor.Invoke(new object[] { rowFirst, rowLast, columnFirst, columnLast, sheetId });
        }

        internal static object CreateExcelReference(int[][] areas, IntPtr sheetId)
        {
            return excelReferenceConstructorRects.Invoke(new object[] { areas, sheetId });
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
        #endregion 

        #region ExcelError
        internal static bool IsExcelErrorObject(object o)
        {
            return excelErrorType.IsInstanceOfType(o);
        }

        internal static int ExcelErrorGetValue(object e)
        {
            return (int)(short)e;
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
        #endregion

        #region ExcelMissing
        internal static bool IsExcelMissingObject(object o)
        {
            return excelMissingType.IsInstanceOfType(o);
        }
        internal static object GetExcelMissingValue()
        {
            return excelMissingValue;
        }
        #endregion

        #region ExcelEmpty
        internal static bool IsExcelEmptyObject(object o)
        {
            return excelEmptyType.IsInstanceOfType(o);
        }

        internal static object GetExcelEmptyValue()
        {
            return excelEmptyValue;
        }
        #endregion

        #region ExcelAsyncHandle
        internal static bool IsExcelAsyncHandleNativeObject(object o)
        {
            return excelAsyncHandleNativeType.IsInstanceOfType(o);
        }

        internal static object CreateExcelAsyncHandleNative(IntPtr handle)
        {
            return excelAsyncHandleNativeConstructor.Invoke(new object[] { handle });
        }

        internal static IntPtr GetExcelAsyncHandleNativeHandle(object o)
        {
            return (IntPtr)excelAsyncHandleNativeHandleField.GetValue(o);
        }

        // We need this for the parameter setup, which has  a special case for this type.
        internal static Type ExcelAsyncHandleType { get { return excelAsyncHandleType; } }
        #endregion
    }
}
