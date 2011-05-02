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
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

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
		public Type BoxedValueType; 	// Causes a wrapper to be created that boxes the return type from the user method,
											// allowing Custom Marshaling to be injected

		public XlParameterInfo(ParameterInfo paramInfo)
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

			// Get Description
			object[] attribs = paramInfo.GetCustomAttributes(false);
			// Search through attribs for Description
			foreach (object attrib in attribs)
			{
				System.ComponentModel.DescriptionAttribute desc =
					attrib as System.ComponentModel.DescriptionAttribute;
				if (desc != null)
				{
					Description = desc.Description;
				}
				//// HACK: Some problem with library references - 
				//// For now relax the assembly reference and use late-bound
                Type attribType = attrib.GetType();
                if (attribType.FullName == "ExcelDna.Integration.ExcelArgumentAttribute")
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
			SetTypeInfo(paramInfo.ParameterType, false, false);
		}

		public XlParameterInfo(Type type, bool isReturnType, bool isExceptionSafe)
		{
			SetTypeInfo(type, isReturnType, isExceptionSafe);
		}

        public void SetTypeInfo(Type type, bool isReturnType, bool isExceptionSafe)
        {
            if (XlAddIn.XlCallVersion < 12)
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
					if (!isExceptionSafe)
					{
						XlType = "P"; // OPER
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
						DelegateParamType = typeof(object);
						BoxedValueType = typeof(bool);
					}
					else
					{
						XlType = "P"; // OPER
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlBooleanReturnMarshaler));
						DelegateParamType = typeof(object);
						BoxedValueType = typeof(bool);
					}
				}
				else
				{
					XlType = "J"; // int32
				}
			}
			else if (type == typeof(int))
			{
				if (isReturnType && !isExceptionSafe)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(int);
				}
				else
				{
					XlType = "J";
				}
			}
			else if (type == typeof(short))
			{
				if (isReturnType && !isExceptionSafe)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(short);
				}
				else
				{
					XlType = "I";
				}
			}
			else if (type == typeof(ushort))
			{
				if (isReturnType && !isExceptionSafe)
				{
					XlType = "P"; // OPER
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObjectMarshaler));
					DelegateParamType = typeof(object);
					BoxedValueType = typeof(ushort);
				}
				else
				{
					XlType = "H";
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
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlLongParameterMarshaler));
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
                        XlType = "Q"; // OPER
                        MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                        DelegateParamType = typeof(object);
                    }
                    else
                    {
                        XlType = "D%"; // XlString12
                        MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlString12ReturnMarshaler));
                    }
                }
                else
                {
                    XlType = "C%"; // LPWSTR
                    MarshalAsAttribute = GetMarshalAsAttribute(UnmanagedType.LPWStr);
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
                    }
                    else
                    {
                        // TODO: Consolidate with the above case?
                        XlType = "E"; // double*
						MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDateTime12Marshaler));
                        DelegateParamType = typeof(object);
                        BoxedValueType = typeof(DateTime);
                    }
                }
                else
                {
					XlType = "E"; // double*
					MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDateTime12Marshaler));
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
                XlType = "K%"; // FP14*
                MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDoubleArray12Marshaler), "1");
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
                    XlType = "U"; // XLOPER
                else
                    XlType = "Q"; // OPER
                MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                DelegateParamType = typeof(object);
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
            }
            else if (type == typeof(bool))
            {
                if (isReturnType)
                {
                    if (!isExceptionSafe)
                    {
                        XlType = "Q"; // OPER
                        MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                        DelegateParamType = typeof(object);
                        BoxedValueType = typeof(bool);
                    }
                    else
                    {
                        XlType = "Q"; // OPER
                        MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlBoolean12ReturnMarshaler));
                        DelegateParamType = typeof(object);
                        BoxedValueType = typeof(bool);
                    }
                }
                else
                {
                    XlType = "J"; // int32
                }
            }
            else if (type == typeof(int))
            {
                if (isReturnType && !isExceptionSafe)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(int);
                }
                else
                {
                    XlType = "J";
                }
            }
            else if (type == typeof(short))
            {
                if (isReturnType && !isExceptionSafe)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(short);
                }
                else
                {
                    XlType = "I";
                }
            }
            else if (type == typeof(ushort))
            {
                if (isReturnType && !isExceptionSafe)
                {
                    XlType = "Q"; // OPER
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlObject12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(ushort);
                }
                else
                {
                    XlType = "H";
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
                }
                else
                {
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlDecimalParameter12Marshaler));
                    DelegateParamType = typeof(object);
                    BoxedValueType = typeof(decimal);
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
                }
                else
                {
                    XlType = "E"; // double*
                    MarshalAsAttribute = GetMarshalAsAttribute(typeof(XlLongParameter12Marshaler));
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
