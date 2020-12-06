//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;

namespace ExcelDna.Loader
{
    // Return type and all parameters are described by this structure.
    internal class XlParameterInfo
	{
		public string XlType;
		public string Name;         // Ignored for return 'parameter'
		public string Description;  // Ignored for return 'parameter'
		public bool AllowReference; // Ignored for return 'parameter'
		public CustomAttributeBuilder MarshalAsAttribute;
		public Type DelegateParamType;
		public Type BoxedValueType;     // Causes a wrapper to be created that boxes the return type from the user method,
                                        // allowing Custom Marshaling to be injected

        // Added for DirectMarshal
        public string DirectMarshalXlType;      // Identifier for Excel Type
        public MethodInfo DirectMarshalConvert; // Method on XlMarshalContext w/ 1 param and return value: Native -> Managed for parameters, and Managed -> Native for Return

        public XlParameterInfo(ParameterInfo paramInfo, object attrib)
		{
			// Add Name and Description
			// CONSIDER: Override Marshaler for row/column arrays according to some attribute

			// Some pre-checks
			if (paramInfo.ParameterType.IsByRef)
				throw new DnaMarshalException("Parameter is ByRef: " + paramInfo.Name);
			
			// Default Name and Description
			Name = paramInfo.Name;
			Description = "";
			AllowReference = false;

            SetAttributeInfo(attrib);
			SetTypeInfo(paramInfo.ParameterType, false, false);
		}

        public XlParameterInfo(Type type, bool isReturnType, bool isExceptionSafe)
        {
            SetTypeInfo(type, isReturnType, isExceptionSafe);
        }

        public bool IsExcelAsyncHandle { get { return XlType == "X"; } }

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
            if ((XlAddIn.XlCallVersion < 12)
//#if DEBUG // on debug-32bit we want to keep exercising the Excel4 APIs
//                || (IntPtr.Size == 4)
//#endif
            )
            {
                SetTypeInfo4(type, isReturnType, isExceptionSafe);
            }
            else
            {
                SetTypeInfo12(type, isReturnType, isExceptionSafe);
            }
        }

        private void SetTypeInfo4(Type type, bool isReturnType, bool isExceptionSafe)
		{
			// isExceptionSafe determines whether or not exception wrapper will be constructed
			// if isExceptionSafe then no exception wrapper is created
			// else the wrapper function returns an object, and the XlObjectMarshaler is always 
            // used - the wrapper then ensures that #ERROR is returned from the function 
			// if any exception is caught.
			// if no exception, the return type is known to be of type BoxedReturnValueType
			// and unboxed accordingly.

            // NOTE: There is also a list of supported parameter types in
            // AssemblyLoaded.cs, where the methods to register are extracted.
            
            // By default DelegateParamType is type
			// changed for some return types to ensure boxing,
			// to allow custom marshaling.
			DelegateParamType = type;

			if (type == typeof(double))
			{
				if (isReturnType && !isExceptionSafe)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(double);
				}
				else
				{
					XlType = "B";
				}
			}
			else if (type == typeof(string))
			{
				// CONSIDER: Other options for string marshaling (nulls etc??)
				if (isReturnType)
				{
					if (!isExceptionSafe)
					{
						XlType = "P"; // OPER
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
						DelegateParamType = typeof(object);
					}
					else
					{
						XlType = "D"; // byte-counted string *
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlStringReturnMarshaler));
					}
				}
				else
				{
					XlType = "C"; // LPSTR
					MarshalAsAttribute = GetMarshalAsAttribute(UnmanagedType.LPStr);
				}
			}
			else if (type == typeof(DateTime))
			{
				if (isReturnType)
				{
					if (!isExceptionSafe)
					{
						XlType = "P"; // OPER
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
						DelegateParamType = typeof(object);
						BoxedValueType = typeof(DateTime);
					}
					else
					{
						// TODO: Consolidate with the above case? - NO! Cluster Connector does not allow OPER types
						XlType = "E"; // double*
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDateTimeMarshaler));
						DelegateParamType = typeof(object);
						BoxedValueType = typeof(DateTime);
					}
				}
				else
				{
					XlType = "E"; // double*
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDateTimeMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(DateTime);
				}
			}
			else if (type == typeof(double[]))
			{
				//if (isReturnType && !isExceptionSafe)
				//{
				//    XlType = 'P'; // OPER
				//    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
				//    DelegateParamType = typeof(object);
				//}
				//else
				//{
					XlType = "K"; // FP*
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDoubleArrayMarshaler), "1");
				//}
			}
			else if (type == typeof(double[,]))
			{
				//if (isReturnType && !isExceptionSafe)
				//{
				//    XlType = 'P'; // OPER
				//    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
				//    DelegateParamType = typeof(object);
				//}
				//else
				//{
					XlType = "K"; // FP*
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDoubleArrayMarshaler), "2");
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

                if (AllowReference)
					XlType = "R"; // XLOPER
				else
					XlType = "P"; // OPER
				MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
				DelegateParamType = typeof(object);
			}
			else if (type == typeof(object[]))
			{
				if (isReturnType && !isExceptionSafe)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
				}
				else
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectArrayMarshaler), "1");
				}
			}
			else if (type == typeof(object[,]))
			{
				if (isReturnType && !isExceptionSafe)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
				}
				else
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectArrayMarshaler), "2");
				}
			}
			else if (type == typeof(bool))
			{
				if (isReturnType)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(bool);

				}
				else
				{
					// XlType = "J"; // int32
                    XlType = "P"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlBooleanMarshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(bool);
				}
			}
			else if (type == typeof(int))
			{
				if (isReturnType)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(int);
				}
				else
				{
					// XlType = "J";
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlInt32ParameterMarshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(int);
				}
			}
			else if (type == typeof(short))
			{
				if (isReturnType)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(short);
				}
				else
				{
					// XlType = "I";
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlInt16ParameterMarshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(short);
				}
			}
			else if (type == typeof(ushort))
			{
				if (isReturnType)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(ushort);
				}
				else
				{
					// XlType = "H";
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlUInt16ParameterMarshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(ushort);
				}
			}
            else if (type == typeof(decimal))
            {
                if (isReturnType)
                {
                    XlType = "P"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(decimal);
                }
                else
                {
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDecimalParameterMarshaler));
				    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(decimal);
                }
            }
            else if (type == typeof(long))
            {
                // Just like decimal - change to double as well as we can.
                if (isReturnType)
                {
                    XlType = "P"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(long);
                }
                else
                {
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlInt64ParameterMarshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(long);
                }
            }
            else
			{
				// The function is bad and cannot be marshaled to Excel
				throw new DnaMarshalException("Unknown Data Type: " + type.ToString());
			}
		}

        private void SetTypeInfo12(Type type, bool isReturnType, bool isExceptionSafe)
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

            // By default DelegateParamType is type
            // changed for some return types to ensure boxing,
            // to allow custom marshaling.
            DelegateParamType = type;

            if (type == typeof(double))
            {
                if (isReturnType && !isExceptionSafe)
                {
                    XlType = "Q"; // OPER12
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(double);
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = XlMarshalConversions.DoubleToXloperReturn;
                }
                else
                {
                    XlType = "B";
                    DirectMarshalXlType = XlTypes.DoublePtr;
                    if (isReturnType)
                        DirectMarshalConvert = XlMarshalConversions.DoublePtrReturn;
                    else
                        DirectMarshalConvert = XlMarshalConversions.DoublePtrParam;
                }
            }
            else if (type == typeof(string))
            {
                // CONSIDER: Other options for string marshaling (nulls etc??)
                if (isReturnType)
                {
                    if (!isExceptionSafe)
                    {
                        XlType = "Q"; // OPER
                        MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                        DelegateParamType = typeof(object);
                        DirectMarshalXlType = XlTypes.Xloper;
                        DirectMarshalConvert = XlMarshalConversions.ObjectReturn;
                    }
                    else
                    {
                        XlType = "D%"; // XlString12
                        MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlString12ReturnMarshaler));
                        DirectMarshalXlType = XlTypes.String;
                        DirectMarshalConvert = XlMarshalConversions.StringReturn;
                    }
                }
                else
                {
                    XlType = "C%"; // LPWSTR
                    MarshalAsAttribute = GetMarshalAsAttribute(UnmanagedType.LPWStr);
                    DirectMarshalXlType = XlTypes.String;
                    DirectMarshalConvert = XlMarshalConversions.StringParam;
                }
            }
            else if (type == typeof(DateTime))
            {
                if (isReturnType)
                {
                    if (!isExceptionSafe)
                    {
                        XlType = "Q"; // OPER
                        MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                        DelegateParamType = typeof(object);
                        BoxedValueType = typeof(DateTime);
                        DirectMarshalXlType = XlTypes.Xloper;
                        DirectMarshalConvert = XlMarshalConversions.DateTimeToXloperReturn;
                    }
                    else
                    {
                        // TODO: Consolidate with the above case?
                        XlType = "E"; // double*
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDateTime12Marshaler));
                        DelegateParamType = typeof(object);
                        BoxedValueType = typeof(DateTime);
                        DirectMarshalXlType = XlTypes.DoublePtr;
                        DirectMarshalConvert = XlMarshalConversions.DateTimeToDoublePtrReturn;
                    }
                }
                else
                {
					XlType = "E"; // double*
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDateTime12Marshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(DateTime);
                    DirectMarshalXlType = XlTypes.DoublePtr;
                    DirectMarshalConvert = XlMarshalConversions.DateTimeFromDoublePtrParam;
                }
            }
            else if (type == typeof(double[]))
            {
                //if (isReturnType && !isExceptionSafe)
                //{
                //    XlType = 'P'; // OPER
                //    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
                //    DelegateParamType = typeof(object);
                //}
                //else
                //{
                XlType = "K%"; // FP12*
                MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDoubleArray12Marshaler), "1");
                DirectMarshalXlType = XlTypes.DoubleArray;
                if (isReturnType)
                    DirectMarshalConvert = XlMarshalConversions.DoubleArray1Return;
                else
                    DirectMarshalConvert = XlMarshalConversions.DoubleArray1Param;
                //}
            }
            else if (type == typeof(double[,]))
            {
                //if (isReturnType && !isExceptionSafe)
                //{
                //    XlType = 'P'; // OPER
                //    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
                //    DelegateParamType = typeof(object);
                //}
                //else
                //{
                XlType = "K%"; // FP12*
                MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDoubleArray12Marshaler), "2");
                DirectMarshalXlType = XlTypes.DoubleArray;
                if (isReturnType)
                    DirectMarshalConvert = XlMarshalConversions.DoubleArray2Return;
                else
                    DirectMarshalConvert = XlMarshalConversions.DoubleArray2Param;
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
                {
                    XlType = "U"; // XLOPER
                    DirectMarshalXlType = XlTypes.XloperAllowRef;
                }
                else
                {
                    XlType = "Q"; // OPER
                    DirectMarshalXlType = XlTypes.Xloper;
                }

                MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                if (isReturnType)
                    DirectMarshalConvert = XlMarshalConversions.ObjectReturn;
                else
                    DirectMarshalConvert = XlMarshalConversions.ObjectParam;
            }
            else if (type == typeof(object[]))
            {
                if (isReturnType && !isExceptionSafe)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                }
                else
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectArray12Marshaler), "1");
                }

                DirectMarshalXlType = XlTypes.Xloper;
                if (isReturnType)
                    DirectMarshalConvert = XlMarshalConversions.ObjectArray1Return;
                else
                    DirectMarshalConvert = XlMarshalConversions.ObjectArray1Param;
            }
            else if (type == typeof(object[,]))
            {
                if (isReturnType && !isExceptionSafe)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                }
                else
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectArray12Marshaler), "2");
                }

                DirectMarshalXlType = XlTypes.Xloper;
                if (isReturnType)
                    DirectMarshalConvert = XlMarshalConversions.ObjectArray2Return;
                else
                    DirectMarshalConvert = XlMarshalConversions.ObjectArray2Param;
            }
            else if (type == typeof(bool))
            {
                if (isReturnType)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(bool);
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = XlMarshalConversions.BoolToXloperReturn;
                }
                else
                {
					// XlType = "J"; // int32
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlBoolean12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(bool);
                    DirectMarshalXlType = XlTypes.BoolPtr;
                    DirectMarshalConvert = XlMarshalConversions.BoolPtrParam;
                }
            }
            else if (type == typeof(int))
            {
                if (isReturnType)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(int);
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = XlMarshalConversions.Int32ToXloperReturn;
                }
                else
                {
                    // XlType = "J";
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlInt32Parameter12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(int);
                    DirectMarshalXlType = XlTypes.DoublePtr;
                    DirectMarshalConvert = XlMarshalConversions.DoublePtrToInt32Param;
                }
            }
            else if (type == typeof(short))
            {
                if (isReturnType)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(short);
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = XlMarshalConversions.Int16ToXloperReturn;
                }
                else
                {
                    // XlType = "I";
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlInt16Parameter12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(short);
                    DirectMarshalXlType = XlTypes.DoublePtr;
                    DirectMarshalConvert = XlMarshalConversions.DoublePtrToInt16Param;
                }
            }
            else if (type == typeof(ushort))
            {
                if (isReturnType)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(ushort);
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = XlMarshalConversions.UInt16ToXloperReturn;
                }
                else
                {
                    // XlType = "H";
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlUInt16Parameter12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(ushort);
                    DirectMarshalXlType = XlTypes.DoublePtr;
                    DirectMarshalConvert = XlMarshalConversions.DoublePtrToUInt16Param;
                }
            }
            else if (type == typeof(decimal))
            {
                if (isReturnType)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(decimal);
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = XlMarshalConversions.DecimalToXloperReturn;
                }
                else
                {
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDecimalParameter12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(decimal);
                    DirectMarshalXlType = XlTypes.DoublePtr;
                    DirectMarshalConvert = XlMarshalConversions.DoublePtrToDecimalParam;
                }
            }
            else if (type == typeof(long))
            {
                if (isReturnType)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(long);
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = XlMarshalConversions.Int64ToXloperReturn;
                }
                else
                {
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlInt64Parameter12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(long);
                    DirectMarshalXlType = XlTypes.DoublePtr;
                    DirectMarshalConvert = XlMarshalConversions.DoublePtrToInt64Param;
                }
            }
            else if (type == IntegrationMarshalHelpers.ExcelAsyncHandleType && !isReturnType)
            {
                XlType = "X"; // Async Handle in XLOPER12's BigData
                MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlAsyncHandleParameter12Marshaler));
                DirectMarshalXlType = XlTypes.AsyncHandle;
                DirectMarshalConvert = XlMarshalConversions.AsyncHandleParam;    // TODO/DM We need a cast here ????
            }
            else if (type == typeof(XlOper12*))
            {
                // Internal use only - the memory management rules are a mess to support generally
                // Direct pointer passthrough, no marshalling
                if (AllowReference)
                {
                    XlType = "U"; // XLOPER
                    DirectMarshalXlType = XlTypes.XloperAllowRef;
                    DirectMarshalConvert = null;
                }
                else
                {
                    XlType = "Q"; // OPER
                    DirectMarshalXlType = XlTypes.Xloper;
                    DirectMarshalConvert = null;
                }
            }
            else
            {
                // The function is bad and cannot be marshaled to Excel
                throw new DnaMarshalException("Unknown Data Type: " + type.ToString());
            }
        }

		private static CustomAttributeBuilder GetMarshalAsAttribute(UnmanagedType unmanagedType)
		{
			Type[] ctorParams = new Type[] { typeof(UnmanagedType) };
			ConstructorInfo classCtorInfo = typeof(MarshalAsAttribute).GetConstructor(ctorParams);

			CustomAttributeBuilder builder = new CustomAttributeBuilder(
								classCtorInfo,
								new object[] { unmanagedType });
			return builder;
		}

		private static CustomAttributeBuilder GetMarshalAsAttribute(Type marshalTypeRef)
		{
			// CONSIDER: Caching some of the metadata loaded here
			Type[] ctorParams = new Type[] { typeof(UnmanagedType) };
			ConstructorInfo classCtorInfo = typeof(MarshalAsAttribute).GetConstructor(ctorParams);

			CustomAttributeBuilder builder = new CustomAttributeBuilder(
								classCtorInfo,
								new object[] { UnmanagedType.CustomMarshaler },
								new FieldInfo[] { typeof(MarshalAsAttribute).GetField("MarshalTypeRef") },
								new object[] { marshalTypeRef });
			return builder;
		}

		private static CustomAttributeBuilder GetMarshalAsAttribute(Type marshalTypeRef, string marshalCookie)
		{
			// CONSIDER: Caching some of the metadata loaded here
			Type[] ctorParams = new Type[] { typeof(UnmanagedType) };
			ConstructorInfo classCtorInfo = typeof(MarshalAsAttribute).GetConstructor(ctorParams);

			CustomAttributeBuilder builder = new CustomAttributeBuilder(
								classCtorInfo,
								new object[] { UnmanagedType.CustomMarshaler },
								new FieldInfo[] { typeof(MarshalAsAttribute).GetField("MarshalTypeRef"), typeof(MarshalAsAttribute).GetField("MarshalCookie") },
								new object[] { marshalTypeRef, marshalCookie });
			return builder;
		}
	}
}
