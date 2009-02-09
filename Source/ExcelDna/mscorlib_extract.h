// This file is an extract from the compiler-generated file:
// C++ source equivalent of Win32 type library C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb

#pragma once
#pragma pack(push, 8)
//
#include <comdef.h>

extern "C" const GUID __declspec(selectany) IID__AppDomain =
    {0x05f696dc,0x2b29,0x3663,{0xad,0x8b,0xc4,0x38,0x9c,0xf2,0xa7,0x13}};
extern "C" const GUID __declspec(selectany) IID__ObjectHandle =
    {0xea675b47,0x64e0,0x3b5f,{0x9b,0xe7,0xf7,0xdc,0x29,0x90,0x73,0x0d}};
extern "C" const GUID __declspec(selectany) IID_IAppDomainSetup =
    {0x27fff232,0xa7a8,0x40dd,{0x8d,0x4a,0x73,0x4a,0xd5,0x9f,0xcd,0x41}};
extern "C" const GUID __declspec(selectany) IID__Type =
    {0xbca8b44d,0xaad6,0x3a86,{0x8a,0xb7,0x03,0x34,0x9f,0x4f,0x2d,0xa2}};
extern "C" const GUID __declspec(selectany) IID__Object =
    {0x65074f7f,0x63c0,0x304e,{0xaf,0x0a,0xd5,0x17,0x41,0xcb,0x4a,0x8d}};

struct __declspec(uuid("65074f7f-63c0-304e-af0a-d51741cb4a8d"))
_Object : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ToString (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall Equals (
        /*[in]*/ VARIANT obj,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetHashCode (
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall GetType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
};

struct __declspec(uuid("bca8b44d-aad6-3a86-8ab7-03349f4f2da2"))
_Type : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetTypeInfoCount (
        /*[out]*/ unsigned long * pcTInfo ) = 0;
      virtual HRESULT __stdcall GetTypeInfo (
        /*[in]*/ unsigned long iTInfo,
        /*[in]*/ unsigned long lcid,
        /*[in]*/ long ppTInfo ) = 0;
      virtual HRESULT __stdcall GetIDsOfNames (
        /*[in]*/ GUID * riid,
        /*[in]*/ long rgszNames,
        /*[in]*/ unsigned long cNames,
        /*[in]*/ unsigned long lcid,
        /*[in]*/ long rgDispId ) = 0;
      virtual HRESULT __stdcall Invoke (
        /*[in]*/ unsigned long dispIdMember,
        /*[in]*/ GUID * riid,
        /*[in]*/ unsigned long lcid,
        /*[in]*/ short wFlags,
        /*[in]*/ long pDispParams,
        /*[in]*/ long pVarResult,
        /*[in]*/ long pExcepInfo,
        /*[in]*/ long puArgErr ) = 0;
      virtual HRESULT __stdcall get_ToString (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall Equals (
        /*[in]*/ VARIANT other,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetHashCode (
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall GetType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_MemberType (
        /*[out,retval]*/ enum MemberTypes * pRetVal ) = 0;
      virtual HRESULT __stdcall get_name (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_DeclaringType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_ReflectedType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetCustomAttributes (
        /*[in]*/ struct _Type * attributeType,
        /*[in]*/ VARIANT_BOOL inherit,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetCustomAttributes_2 (
        /*[in]*/ VARIANT_BOOL inherit,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall IsDefined (
        /*[in]*/ struct _Type * attributeType,
        /*[in]*/ VARIANT_BOOL inherit,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Guid (
        /*[out,retval]*/ GUID * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Module (
        /*[out,retval]*/ struct _Module * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Assembly (
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_TypeHandle (
        /*[out,retval]*/ struct RuntimeTypeHandle * pRetVal ) = 0;
      virtual HRESULT __stdcall get_FullName (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Namespace (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_AssemblyQualifiedName (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall GetArrayRank (
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall get_BaseType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetConstructors (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetInterface (
        /*[in]*/ BSTR name,
        /*[in]*/ VARIANT_BOOL ignoreCase,
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetInterfaces (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall FindInterfaces (
        /*[in]*/ struct _TypeFilter * filter,
        /*[in]*/ VARIANT filterCriteria,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetEvent (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ struct _EventInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetEvents (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetEvents_2 (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetNestedTypes (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetNestedType (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMember (
        /*[in]*/ BSTR name,
        /*[in]*/ enum MemberTypes Type,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetDefaultMembers (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall FindMembers (
        /*[in]*/ enum MemberTypes MemberType,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _MemberFilter * filter,
        /*[in]*/ VARIANT filterCriteria,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetElementType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall IsSubclassOf (
        /*[in]*/ struct _Type * c,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall IsInstanceOfType (
        /*[in]*/ VARIANT o,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall IsAssignableFrom (
        /*[in]*/ struct _Type * c,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetInterfaceMap (
        /*[in]*/ struct _Type * interfaceType,
        /*[out,retval]*/ struct InterfaceMapping * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethod (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ SAFEARRAY * types,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[out,retval]*/ struct _MethodInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethod_2 (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ struct _MethodInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethods (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetField (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ struct _FieldInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetFields (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperty (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ struct _PropertyInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperty_2 (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ struct _Type * returnType,
        /*[in]*/ SAFEARRAY * types,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[out,retval]*/ struct _PropertyInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperties (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMember_2 (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMembers (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall InvokeMember (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags invokeAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ VARIANT Target,
        /*[in]*/ SAFEARRAY * args,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[in]*/ struct _CultureInfo * culture,
        /*[in]*/ SAFEARRAY * namedParameters,
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall get_UnderlyingSystemType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall InvokeMember_2 (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags invokeAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ VARIANT Target,
        /*[in]*/ SAFEARRAY * args,
        /*[in]*/ struct _CultureInfo * culture,
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall InvokeMember_3 (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags invokeAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ VARIANT Target,
        /*[in]*/ SAFEARRAY * args,
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall GetConstructor (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ enum CallingConventions callConvention,
        /*[in]*/ SAFEARRAY * types,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[out,retval]*/ struct _ConstructorInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetConstructor_2 (
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ SAFEARRAY * types,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[out,retval]*/ struct _ConstructorInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetConstructor_3 (
        /*[in]*/ SAFEARRAY * types,
        /*[out,retval]*/ struct _ConstructorInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetConstructors_2 (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_TypeInitializer (
        /*[out,retval]*/ struct _ConstructorInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethod_3 (
        /*[in]*/ BSTR name,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ enum CallingConventions callConvention,
        /*[in]*/ SAFEARRAY * types,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[out,retval]*/ struct _MethodInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethod_4 (
        /*[in]*/ BSTR name,
        /*[in]*/ SAFEARRAY * types,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[out,retval]*/ struct _MethodInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethod_5 (
        /*[in]*/ BSTR name,
        /*[in]*/ SAFEARRAY * types,
        /*[out,retval]*/ struct _MethodInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethod_6 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _MethodInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMethods_2 (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetField_2 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _FieldInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetFields_2 (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetInterface_2 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetEvent_2 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _EventInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperty_3 (
        /*[in]*/ BSTR name,
        /*[in]*/ struct _Type * returnType,
        /*[in]*/ SAFEARRAY * types,
        /*[in]*/ SAFEARRAY * modifiers,
        /*[out,retval]*/ struct _PropertyInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperty_4 (
        /*[in]*/ BSTR name,
        /*[in]*/ struct _Type * returnType,
        /*[in]*/ SAFEARRAY * types,
        /*[out,retval]*/ struct _PropertyInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperty_5 (
        /*[in]*/ BSTR name,
        /*[in]*/ SAFEARRAY * types,
        /*[out,retval]*/ struct _PropertyInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperty_6 (
        /*[in]*/ BSTR name,
        /*[in]*/ struct _Type * returnType,
        /*[out,retval]*/ struct _PropertyInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperty_7 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _PropertyInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetProperties_2 (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetNestedTypes_2 (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetNestedType_2 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMember_3 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetMembers_2 (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Attributes (
        /*[out,retval]*/ enum TypeAttributes * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsNotPublic (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsPublic (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsNestedPublic (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsNestedPrivate (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsNestedFamily (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsNestedAssembly (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsNestedFamANDAssem (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsNestedFamORAssem (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsAutoLayout (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsLayoutSequential (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsExplicitLayout (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsClass (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsInterface (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsValueType (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsAbstract (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsSealed (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsEnum (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsSpecialName (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsImport (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsSerializable (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsAnsiClass (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsUnicodeClass (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsAutoClass (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsArray (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsByRef (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsPointer (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsPrimitive (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsCOMObject (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_HasElementType (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsContextful (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall get_IsMarshalByRef (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall Equals_2 (
        /*[in]*/ struct _Type * o,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
};

enum __declspec(uuid("3223e024-5d70-3236-a92a-6b4114b2632f"))
BindingFlags
{
    BindingFlags_Default = 0,
    BindingFlags_IgnoreCase = 1,
    BindingFlags_DeclaredOnly = 2,
    BindingFlags_Instance = 4,
    BindingFlags_Static = 8,
    BindingFlags_Public = 16,
    BindingFlags_NonPublic = 32,
    BindingFlags_FlattenHierarchy = 64,
    BindingFlags_InvokeMethod = 256,
    BindingFlags_CreateInstance = 512,
    BindingFlags_GetField = 1024,
    BindingFlags_SetField = 2048,
    BindingFlags_GetProperty = 4096,
    BindingFlags_SetProperty = 8192,
    BindingFlags_PutDispProperty = 16384,
    BindingFlags_PutRefDispProperty = 32768,
    BindingFlags_ExactBinding = 65536,
    BindingFlags_SuppressChangeType = 131072,
    BindingFlags_OptionalParamBinding = 262144,
    BindingFlags_IgnoreReturn = 16777216
};

struct __declspec(uuid("27fff232-a7a8-40dd-8d4a-734ad59fcd41"))
IAppDomainSetup : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ApplicationBase (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_ApplicationBase (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_ApplicationName (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_ApplicationName (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_CachePath (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_CachePath (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_ConfigurationFile (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_ConfigurationFile (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_DynamicBase (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_DynamicBase (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_LicenseFile (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_LicenseFile (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_PrivateBinPath (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_PrivateBinPath (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_PrivateBinPathProbe (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_PrivateBinPathProbe (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_ShadowCopyDirectories (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_ShadowCopyDirectories (
        /*[in]*/ BSTR pRetVal ) = 0;
      virtual HRESULT __stdcall get_ShadowCopyFiles (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall put_ShadowCopyFiles (
        /*[in]*/ BSTR pRetVal ) = 0;
};

struct __declspec(uuid("ea675b47-64e0-3b5f-9be7-f7dc2990730d"))
/* dual interface */ _ObjectHandle;

struct __declspec(uuid("ea675b47-64e0-3b5f-9be7-f7dc2990730d"))
_ObjectHandle : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ToString (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall Equals (
        /*[in]*/ VARIANT obj,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetHashCode (
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall GetType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetLifetimeService (
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall InitializeLifetimeService (
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateObjRef (
        /*[in]*/ struct _Type * requestedType,
        /*[out,retval]*/ struct _ObjRef * * pRetVal ) = 0;
      virtual HRESULT __stdcall Unwrap (
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
};

struct __declspec(uuid("05f696dc-2b29-3663-ad8b-c4389cf2a713"))
/* dual interface */ _AppDomain;

struct __declspec(uuid("05f696dc-2b29-3663-ad8b-c4389cf2a713"))
_AppDomain : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ToString (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall Equals (
        /*[in]*/ VARIANT other,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetHashCode (
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall GetType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall InitializeLifetimeService (
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall GetLifetimeService (
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Evidence (
        /*[out,retval]*/ struct _Evidence * * pRetVal ) = 0;
      virtual HRESULT __stdcall add_DomainUnload (
        /*[in]*/ struct _EventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_DomainUnload (
        /*[in]*/ struct _EventHandler * value ) = 0;
      virtual HRESULT __stdcall add_AssemblyLoad (
        /*[in]*/ struct _AssemblyLoadEventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_AssemblyLoad (
        /*[in]*/ struct _AssemblyLoadEventHandler * value ) = 0;
      virtual HRESULT __stdcall add_ProcessExit (
        /*[in]*/ struct _EventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_ProcessExit (
        /*[in]*/ struct _EventHandler * value ) = 0;
      virtual HRESULT __stdcall add_TypeResolve (
        /*[in]*/ struct _ResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_TypeResolve (
        /*[in]*/ struct _ResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall add_ResourceResolve (
        /*[in]*/ struct _ResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_ResourceResolve (
        /*[in]*/ struct _ResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall add_AssemblyResolve (
        /*[in]*/ struct _ResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_AssemblyResolve (
        /*[in]*/ struct _ResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall add_UnhandledException (
        /*[in]*/ struct _UnhandledExceptionEventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_UnhandledException (
        /*[in]*/ struct _UnhandledExceptionEventHandler * value ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_2 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ BSTR dir,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_3 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ struct _Evidence * Evidence,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_4 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ struct _PermissionSet * requiredPermissions,
        /*[in]*/ struct _PermissionSet * optionalPermissions,
        /*[in]*/ struct _PermissionSet * refusedPermissions,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_5 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ BSTR dir,
        /*[in]*/ struct _Evidence * Evidence,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_6 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ BSTR dir,
        /*[in]*/ struct _PermissionSet * requiredPermissions,
        /*[in]*/ struct _PermissionSet * optionalPermissions,
        /*[in]*/ struct _PermissionSet * refusedPermissions,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_7 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ struct _Evidence * Evidence,
        /*[in]*/ struct _PermissionSet * requiredPermissions,
        /*[in]*/ struct _PermissionSet * optionalPermissions,
        /*[in]*/ struct _PermissionSet * refusedPermissions,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_8 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ BSTR dir,
        /*[in]*/ struct _Evidence * Evidence,
        /*[in]*/ struct _PermissionSet * requiredPermissions,
        /*[in]*/ struct _PermissionSet * optionalPermissions,
        /*[in]*/ struct _PermissionSet * refusedPermissions,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall DefineDynamicAssembly_9 (
        /*[in]*/ struct _AssemblyName * name,
        /*[in]*/ enum AssemblyBuilderAccess access,
        /*[in]*/ BSTR dir,
        /*[in]*/ struct _Evidence * Evidence,
        /*[in]*/ struct _PermissionSet * requiredPermissions,
        /*[in]*/ struct _PermissionSet * optionalPermissions,
        /*[in]*/ struct _PermissionSet * refusedPermissions,
        /*[in]*/ VARIANT_BOOL IsSynchronized,
        /*[out,retval]*/ struct _AssemblyBuilder * * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstance (
        /*[in]*/ BSTR AssemblyName,
        /*[in]*/ BSTR typeName,
        /*[out,retval]*/ struct _ObjectHandle * * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstanceFrom (
        /*[in]*/ BSTR assemblyFile,
        /*[in]*/ BSTR typeName,
        /*[out,retval]*/ struct _ObjectHandle * * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstance_2 (
        /*[in]*/ BSTR AssemblyName,
        /*[in]*/ BSTR typeName,
        /*[in]*/ SAFEARRAY * activationAttributes,
        /*[out,retval]*/ struct _ObjectHandle * * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstanceFrom_2 (
        /*[in]*/ BSTR assemblyFile,
        /*[in]*/ BSTR typeName,
        /*[in]*/ SAFEARRAY * activationAttributes,
        /*[out,retval]*/ struct _ObjectHandle * * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstance_3 (
        /*[in]*/ BSTR AssemblyName,
        /*[in]*/ BSTR typeName,
        /*[in]*/ VARIANT_BOOL ignoreCase,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ SAFEARRAY * args,
        /*[in]*/ struct _CultureInfo * culture,
        /*[in]*/ SAFEARRAY * activationAttributes,
        /*[in]*/ struct _Evidence * securityAttributes,
        /*[out,retval]*/ struct _ObjectHandle * * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstanceFrom_3 (
        /*[in]*/ BSTR assemblyFile,
        /*[in]*/ BSTR typeName,
        /*[in]*/ VARIANT_BOOL ignoreCase,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ SAFEARRAY * args,
        /*[in]*/ struct _CultureInfo * culture,
        /*[in]*/ SAFEARRAY * activationAttributes,
        /*[in]*/ struct _Evidence * securityAttributes,
        /*[out,retval]*/ struct _ObjectHandle * * pRetVal ) = 0;
      virtual HRESULT __stdcall Load (
        /*[in]*/ struct _AssemblyName * assemblyRef,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall Load_2 (
        /*[in]*/ BSTR assemblyString,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall Load_3 (
        /*[in]*/ SAFEARRAY * rawAssembly,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall Load_4 (
        /*[in]*/ SAFEARRAY * rawAssembly,
        /*[in]*/ SAFEARRAY * rawSymbolStore,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall Load_5 (
        /*[in]*/ SAFEARRAY * rawAssembly,
        /*[in]*/ SAFEARRAY * rawSymbolStore,
        /*[in]*/ struct _Evidence * securityEvidence,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall Load_6 (
        /*[in]*/ struct _AssemblyName * assemblyRef,
        /*[in]*/ struct _Evidence * assemblySecurity,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall Load_7 (
        /*[in]*/ BSTR assemblyString,
        /*[in]*/ struct _Evidence * assemblySecurity,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall ExecuteAssembly (
        /*[in]*/ BSTR assemblyFile,
        /*[in]*/ struct _Evidence * assemblySecurity,
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall ExecuteAssembly_2 (
        /*[in]*/ BSTR assemblyFile,
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall ExecuteAssembly_3 (
        /*[in]*/ BSTR assemblyFile,
        /*[in]*/ struct _Evidence * assemblySecurity,
        /*[in]*/ SAFEARRAY * args,
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall get_FriendlyName (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_BaseDirectory (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_RelativeSearchPath (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_ShadowCopyFiles (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetAssemblies (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall AppendPrivatePath (
        /*[in]*/ BSTR Path ) = 0;
      virtual HRESULT __stdcall ClearPrivatePath ( ) = 0;
      virtual HRESULT __stdcall SetShadowCopyPath (
        /*[in]*/ BSTR s ) = 0;
      virtual HRESULT __stdcall ClearShadowCopyPath ( ) = 0;
      virtual HRESULT __stdcall SetCachePath (
        /*[in]*/ BSTR s ) = 0;
      virtual HRESULT __stdcall SetData (
        /*[in]*/ BSTR name,
        /*[in]*/ VARIANT data ) = 0;
      virtual HRESULT __stdcall GetData (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall SetAppDomainPolicy (
        /*[in]*/ struct _PolicyLevel * domainPolicy ) = 0;
      virtual HRESULT __stdcall SetThreadPrincipal (
        /*[in]*/ struct IPrincipal * principal ) = 0;
      virtual HRESULT __stdcall SetPrincipalPolicy (
        /*[in]*/ enum PrincipalPolicy policy ) = 0;
      virtual HRESULT __stdcall DoCallBack (
        /*[in]*/ struct _CrossAppDomainDelegate * theDelegate ) = 0;
      virtual HRESULT __stdcall get_DynamicDirectory (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
};

struct __declspec(uuid("17156360-2f1a-384a-bc52-fde93c215c5b"))
_Assembly : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ToString (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall Equals (
        /*[in]*/ VARIANT other,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetHashCode (
        /*[out,retval]*/ long * pRetVal ) = 0;
      virtual HRESULT __stdcall GetType (
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_CodeBase (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_EscapedCodeBase (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall GetName (
        /*[out,retval]*/ struct _AssemblyName * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetName_2 (
        /*[in]*/ VARIANT_BOOL copiedName,
        /*[out,retval]*/ struct _AssemblyName * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_FullName (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_EntryPoint (
        /*[out,retval]*/ struct _MethodInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetType_2 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetType_3 (
        /*[in]*/ BSTR name,
        /*[in]*/ VARIANT_BOOL throwOnError,
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetExportedTypes (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetTypes (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetManifestResourceStream (
        /*[in]*/ struct _Type * Type,
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _Stream * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetManifestResourceStream_2 (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _Stream * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetFile (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _FileStream * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetFiles (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetFiles_2 (
        /*[in]*/ VARIANT_BOOL getResourceModules,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetManifestResourceNames (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetManifestResourceInfo (
        /*[in]*/ BSTR resourceName,
        /*[out,retval]*/ struct _ManifestResourceInfo * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Location (
        /*[out,retval]*/ BSTR * pRetVal ) = 0;
      virtual HRESULT __stdcall get_Evidence (
        /*[out,retval]*/ struct _Evidence * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetCustomAttributes (
        /*[in]*/ struct _Type * attributeType,
        /*[in]*/ VARIANT_BOOL inherit,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetCustomAttributes_2 (
        /*[in]*/ VARIANT_BOOL inherit,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall IsDefined (
        /*[in]*/ struct _Type * attributeType,
        /*[in]*/ VARIANT_BOOL inherit,
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
      virtual HRESULT __stdcall GetObjectData (
        /*[in]*/ struct _SerializationInfo * info,
        /*[in]*/ struct StreamingContext Context ) = 0;
      virtual HRESULT __stdcall add_ModuleResolve (
        /*[in]*/ struct _ModuleResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall remove_ModuleResolve (
        /*[in]*/ struct _ModuleResolveEventHandler * value ) = 0;
      virtual HRESULT __stdcall GetType_4 (
        /*[in]*/ BSTR name,
        /*[in]*/ VARIANT_BOOL throwOnError,
        /*[in]*/ VARIANT_BOOL ignoreCase,
        /*[out,retval]*/ struct _Type * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetSatelliteAssembly (
        /*[in]*/ struct _CultureInfo * culture,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetSatelliteAssembly_2 (
        /*[in]*/ struct _CultureInfo * culture,
        /*[in]*/ struct _Version * Version,
        /*[out,retval]*/ struct _Assembly * * pRetVal ) = 0;
      virtual HRESULT __stdcall LoadModule (
        /*[in]*/ BSTR moduleName,
        /*[in]*/ SAFEARRAY * rawModule,
        /*[out,retval]*/ struct _Module * * pRetVal ) = 0;
      virtual HRESULT __stdcall LoadModule_2 (
        /*[in]*/ BSTR moduleName,
        /*[in]*/ SAFEARRAY * rawModule,
        /*[in]*/ SAFEARRAY * rawSymbolStore,
        /*[out,retval]*/ struct _Module * * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstance (
        /*[in]*/ BSTR typeName,
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstance_2 (
        /*[in]*/ BSTR typeName,
        /*[in]*/ VARIANT_BOOL ignoreCase,
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall CreateInstance_3 (
        /*[in]*/ BSTR typeName,
        /*[in]*/ VARIANT_BOOL ignoreCase,
        /*[in]*/ enum BindingFlags bindingAttr,
        /*[in]*/ struct _Binder * Binder,
        /*[in]*/ SAFEARRAY * args,
        /*[in]*/ struct _CultureInfo * culture,
        /*[in]*/ SAFEARRAY * activationAttributes,
        /*[out,retval]*/ VARIANT * pRetVal ) = 0;
      virtual HRESULT __stdcall GetLoadedModules (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetLoadedModules_2 (
        /*[in]*/ VARIANT_BOOL getResourceModules,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetModules (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetModules_2 (
        /*[in]*/ VARIANT_BOOL getResourceModules,
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetModule (
        /*[in]*/ BSTR name,
        /*[out,retval]*/ struct _Module * * pRetVal ) = 0;
      virtual HRESULT __stdcall GetReferencedAssemblies (
        /*[out,retval]*/ SAFEARRAY * * pRetVal ) = 0;
      virtual HRESULT __stdcall get_GlobalAssemblyCache (
        /*[out,retval]*/ VARIANT_BOOL * pRetVal ) = 0;
};

#pragma pack(pop)
