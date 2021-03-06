//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Reflection;

namespace ExcelDna.Loader
{
    // Return type and all parameters are described by this structure.
    internal class XlParameterInfo
	{
		public string Name;         // Ignored for return 'parameter'
		public string Description;  // Ignored for return 'parameter'
		public bool AllowReference; // Ignored for return 'parameter'
        // Marshaling info
        public string XlType;       // Identifier for Excel Type
        public MethodInfo XlMarshalConvert; // Method on XlMarshalContext w/ 1 param and return value: Native -> Managed for parameters, and Managed -> Native for Return

        public XlParameterInfo(string paramName, Type paramType, object attrib)
		{
			// Add Name and Description
			// CONSIDER: Override Marshaler for row/column arrays according to some attribute

			// Some pre-checks
			if (paramType.IsByRef)
				throw new DnaMarshalException("Parameter is ByRef: " + paramName);
			
			// Default Name and Description
			Name = paramName;
			Description = "";
			AllowReference = false;

            SetAttributeInfo(attrib);
			SetTypeInfo(paramType, false, false);
		}

        public XlParameterInfo(Type type, bool isReturnType, bool isExceptionSafe)
        {
            SetTypeInfo(type, isReturnType, isExceptionSafe);
        }

        public bool IsExcelAsyncHandle => XlType == XlTypes.AsyncHandle;

        void SetAttributeInfo(object attrib)
        {
            if (attrib == null) return;

            // Search through attribs for Description
            System.ComponentModel.DescriptionAttribute desc =
                attrib as System.ComponentModel.DescriptionAttribute;
            if (desc != null)
            {
                Description = desc.Description;
                return;
            }
            //// HACK: Some problem with library references - 
            //// For now relax the assembly reference and use late-bound
            Type attribType = attrib.GetType();
            if (TypeHelper.TypeHasAncestorWithFullName(attribType, "ExcelDna.Integration.ExcelArgumentAttribute"))
            {
                string name = (string)attribType.GetField("Name").GetValue(attrib);
                string description = (string)attribType.GetField("Description").GetValue(attrib);
                object allowReference = attribType.GetField("AllowReference").GetValue(attrib);

                if (name != null)
                    Name = name;
                if (description != null)
                    Description = description;
                if (allowReference != null)
                    AllowReference = (bool)allowReference;
            }
            // HACK: Here is the other code:
            //ExcelArgumentAttribute xlparam = attrib as ExcelArgumentAttribute;
            //if (xlparam != null)
            //{
            //    if (xlparam.Name != null)
            //    {
            //        Name = xlparam.Name;
            //    }
            //    if (xlparam.Description != null)
            //    {
            //        Description = xlparam.Description;
            //    }
            //    AllowReference = xlparam.AllowReference;
            //}
        }

        public void SetTypeInfo(Type type, bool isReturnType, bool isExceptionSafe)
        {
            // isExceptionSafe determines whether or not exception wrapper will be constructed
            // if isExceptionSafe then no exception wrapper is created
            // else the wrapper function returns an object, and the XlObjectMarshaler is always 
            // used - the wrapper then ensures that #ERROR is returned from the function 
            // if any exception is caught.
            // if no exception, the return type is known to be of type BoxedReturnValueType
            // and unboxed accordingly.

            // NOTE: There is also a list of supported parameter types in
            // AssemblyLoader.cs, where the methods to register are extracted.

            if (type == typeof(double))
            {
                if (isReturnType && !isExceptionSafe)
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = XlMarshalConversions.DoubleToXloperReturn;
                }
                else
                {
                    XlType = XlTypes.DoublePtr;
                    if (isReturnType)
                        XlMarshalConvert = XlMarshalConversions.DoublePtrReturn;
                    else
                        XlMarshalConvert = XlMarshalConversions.DoublePtrParam;
                }
            }
            else if (type == typeof(string))
            {
                // CONSIDER: Other options for string marshaling (nulls etc??)
                if (isReturnType)
                {
                    if (!isExceptionSafe)
                    {
                        XlType = XlTypes.Xloper;
                        XlMarshalConvert = XlMarshalConversions.ObjectReturn;
                    }
                    else
                    {
                        XlType = XlTypes.String;
                        XlMarshalConvert = XlMarshalConversions.StringReturn;
                    }
                }
                else
                {
                    XlType = XlTypes.String;
                    XlMarshalConvert = XlMarshalConversions.StringParam;
                }
            }
            else if (type == typeof(DateTime))
            {
                if (isReturnType)
                {
                    if (!isExceptionSafe)
                    {
                        XlType = XlTypes.Xloper;
                        XlMarshalConvert = XlMarshalConversions.DateTimeToXloperReturn;
                    }
                    else
                    {
                        XlType = XlTypes.DoublePtr;
                        XlMarshalConvert = XlMarshalConversions.DateTimeToDoublePtrReturn;
                    }
                }
                else
                {
                    XlType = XlTypes.DoublePtr;
                    XlMarshalConvert = XlMarshalConversions.DateTimeFromDoublePtrParam;
                }
            }
            else if (type == typeof(double[]))
            {
                XlType = XlTypes.DoubleArray;
                if (isReturnType)
                    XlMarshalConvert = XlMarshalConversions.DoubleArray1Return;
                else
                    XlMarshalConvert = XlMarshalConversions.DoubleArray1Param;
            }
            else if (type == typeof(double[,]))
            {
                XlType = XlTypes.DoubleArray;
                if (isReturnType)
                    XlMarshalConvert = XlMarshalConversions.DoubleArray2Return;
                else
                    XlMarshalConvert = XlMarshalConversions.DoubleArray2Param;
                //}
            }
            else if (type == typeof(object))
            {
                // Before version 0.29 we had:
                //    if (isReturnType || AllowReference)
                //        XlType = "U"; // XLOPER
                // and thus registered as U in most cases. 
                // This does not work in HPC setting, and seems to have been a mistake anyway 
                // - returning a reference always gives #VALUE
                // 2020: Not true - it now seems fine (and useful) to return an ExcelReference, and that works even if we declared the return value as Q

                if (AllowReference)
                    XlType = XlTypes.XloperAllowRef;
                else
                    XlType = XlTypes.Xloper;

                if (isReturnType)
                    XlMarshalConvert = XlMarshalConversions.ObjectReturn;
                else
                    XlMarshalConvert = XlMarshalConversions.ObjectParam;
            }
            else if (type == typeof(object[]))
            {
                XlType = XlTypes.Xloper;
                if (isReturnType)
                    XlMarshalConvert = XlMarshalConversions.ObjectArray1Return;
                else
                    XlMarshalConvert = XlMarshalConversions.ObjectArray1Param;
            }
            else if (type == typeof(object[,]))
            {
                XlType = XlTypes.Xloper;
                if (isReturnType)
                    XlMarshalConvert = XlMarshalConversions.ObjectArray2Return;
                else
                    XlMarshalConvert = XlMarshalConversions.ObjectArray2Param;
            }
            else if (type == typeof(bool))
            {
                if (isReturnType)
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = XlMarshalConversions.BoolToXloperReturn;
                }
                else
                {
                    XlType = XlTypes.BoolPtr;
                    XlMarshalConvert = XlMarshalConversions.BoolPtrParam;
                }
            }
            else if (type == typeof(int))
            {
                if (isReturnType)
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = XlMarshalConversions.Int32ToXloperReturn;
                }
                else
                {
                    XlType = XlTypes.DoublePtr;
                    XlMarshalConvert = XlMarshalConversions.DoublePtrToInt32Param;
                }
            }
            else if (type == typeof(short))
            {
                if (isReturnType)
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = XlMarshalConversions.Int16ToXloperReturn;
                }
                else
                {
                    XlType = XlTypes.DoublePtr;
                    XlMarshalConvert = XlMarshalConversions.DoublePtrToInt16Param;
                }
            }
            else if (type == typeof(ushort))
            {
                if (isReturnType)
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = XlMarshalConversions.UInt16ToXloperReturn;
                }
                else
                {
                    XlType = XlTypes.DoublePtr;
                    XlMarshalConvert = XlMarshalConversions.DoublePtrToUInt16Param;
                }
            }
            else if (type == typeof(decimal))
            {
                if (isReturnType)
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = XlMarshalConversions.DecimalToXloperReturn;
                }
                else
                {
                    XlType = XlTypes.DoublePtr;
                    XlMarshalConvert = XlMarshalConversions.DoublePtrToDecimalParam;
                }
            }
            else if (type == typeof(long))
            {
                if (isReturnType)
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = XlMarshalConversions.Int64ToXloperReturn;
                }
                else
                {
                    XlType = XlTypes.DoublePtr;
                    XlMarshalConvert = XlMarshalConversions.DoublePtrToInt64Param;
                }
            }
            else if (type == typeof(ExcelDna.Integration.ExcelAsyncHandle) && !isReturnType)
            {
                XlType = XlTypes.AsyncHandle;
                XlMarshalConvert = XlMarshalConversions.AsyncHandleParam;    // TODO/DM We need a cast here ????
            }
            else if (type == typeof(XlOper12*))
            {
                // Internal use only - the memory management rules are a mess to support generally
                // Direct pointer passthrough, no marshalling
                if (AllowReference)
                {
                    XlType = XlTypes.XloperAllowRef;
                    XlMarshalConvert = null;
                }
                else
                {
                    XlType = XlTypes.Xloper;
                    XlMarshalConvert = null;
                }
            }
            else
            {
                // The function is bad and cannot be marshaled to Excel
                throw new DnaMarshalException("Unknown Data Type: " + type.ToString());
            }
        }
	}
}
