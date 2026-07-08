using Microsoft.CodeAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace ExcelDna.SourceGenerator.NativeAOT
{
    public class Util
    {
        public static string GetFullTypeName(ITypeSymbol type)
        {
            return type.ToDisplayString().Replace("[*,*]", "[,]");
        }

        public static string GetFullTypeOfTypeName(ITypeSymbol type)
        {
            return type.ToDisplayString(FullTypeOfTypeNameFormat).Replace("[*,*]", "[,]");
        }

        public static string GetFullGenericTypeName(INamedTypeSymbol type)
        {
            return type.ToDisplayString(FullGenericNameFormat);
        }

        public static string GetFullMethodName(IMethodSymbol method)
        {
            return $"{GetFullTypeName(method.ContainingType)}.{method.Name}";
        }

        public static IEnumerable<ITypeSymbol> GetBaseTypesAndThis(ITypeSymbol type)
        {
            var current = type;
            while (current != null)
            {
                yield return current;
                current = current.BaseType;
            }
        }

        public static bool TypeHasAncestorWithFullName(ITypeSymbol type, string fullName)
        {
            return GetBaseTypesAndThis(type).Any(i => i.ToDisplayString(FullNameFormat) == fullName);
        }

        public static string MethodType(IMethodSymbol method)
        {
            return BuildMethodType(method.Parameters.Select(p => GetFullTypeName(p.Type)), method.ReturnType, method.ReturnsVoid);
        }

        public static string ExtendedMethodType(IMethodSymbol method)
        {
            return BuildExtendedMethodType(method.Parameters.Select(p => GetFullTypeName(p.Type)), method.ReturnType, method.ReturnsVoid);
        }

        public static string MethodPostParameterConversionType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), method.ReturnType, method.ReturnsVoid);
        }

        public static string ExtendedMethodPostParameterConversionType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildExtendedMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), method.ReturnType, method.ReturnsVoid);
        }

        public static string MethodPostParameterConversionType(IMethodSymbol method, string returnTypeName, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), returnTypeName, returnsVoid: false);
        }

        public static string ExtendedMethodPostParameterConversionType(IMethodSymbol method, string returnTypeName, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildExtendedMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), returnTypeName, returnsVoid: false);
        }

        public static string MethodPostReturnConversionType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<IMethodSymbol>? userReturnConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), GetPostReturnConversionReturnTypeName(method, userReturnConversions), returnsVoid: false);
        }

        public static string ExtendedMethodPostReturnConversionType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<IMethodSymbol>? userReturnConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildExtendedMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), GetPostReturnConversionReturnTypeName(method, userReturnConversions), returnsVoid: false);
        }

        public static string AsyncWrapperMethodType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildMethodType(GetAsyncWrapperParameterTypeNames(method, userParameterConversions, excelHandleExternalTypeNames), "object", returnsVoid: false);
        }

        public static string ExtendedAsyncWrapperMethodType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildExtendedMethodType(GetAsyncWrapperParameterTypeNames(method, userParameterConversions, excelHandleExternalTypeNames), "object", returnsVoid: false);
        }

        public static int AsyncWrapperParameterCount(IMethodSymbol method)
        {
            return method.Parameters.Length > 0 && IsCancellationToken(method.Parameters.Last())
                ? method.Parameters.Length - 1
                : method.Parameters.Length;
        }

        public static string AsyncObjectHandleAdapterMethodType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), GetAsyncObjectHandleAdapterReturnTypeName(method), returnsVoid: false);
        }

        public static string ExtendedAsyncObjectHandleAdapterMethodType(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return BuildExtendedMethodType(method.Parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames)), GetAsyncObjectHandleAdapterReturnTypeName(method), returnsVoid: false);
        }

        public static string MethodExpression(string method)
        {
            return $"System.Linq.Expressions.Expression<{method}>";
        }

        public static bool HasPostParameterConversionShape(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return MethodPostParameterConversionType(method, userParameterConversions, excelHandleExternalTypeNames) != MethodType(method);
        }

        public static bool HasPostReturnConversionShape(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions = null, IEnumerable<IMethodSymbol>? userReturnConversions = null, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return !method.ReturnsVoid &&
                   MethodPostReturnConversionType(method, userParameterConversions, userReturnConversions, excelHandleExternalTypeNames) != MethodPostParameterConversionType(method, userParameterConversions, excelHandleExternalTypeNames);
        }

        public static bool IsLastArrayParams(IMethodSymbol method)
        {
            return method.Parameters.Length > 0 && method.Parameters.Last().IsParams && method.Parameters.Last().Type is IArrayTypeSymbol;
        }

        public static bool IsAsyncRegistration(IMethodSymbol method)
        {
            return IsTask(method.ReturnType) ||
                   HasCustomAttribute(method, "ExcelDna.Registration.ExcelAsyncFunctionAttribute");
        }

        public static bool IsObservableRegistration(IMethodSymbol method)
        {
            return IsObservable(method.ReturnType);
        }

        public static bool HasExcelHandleReturn(IMethodSymbol method, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            if (HasReturnCustomAttribute(method, "ExcelDna.Integration.ExcelHandleAttribute"))
                return true;

            ITypeSymbol returnType = method.ReturnType;
            if ((IsTask(returnType) || IsObservable(returnType)) &&
                returnType is INamedTypeSymbol namedReturnType &&
                namedReturnType.TypeArguments.Length == 1)
            {
                returnType = namedReturnType.TypeArguments[0];
            }

            return TypeHasExcelHandle(returnType, excelHandleExternalTypeNames);
        }

        public static bool HasAsyncObjectHandleAdapter(IMethodSymbol method, IEnumerable<string>? excelHandleExternalTypeNames = null)
        {
            return (IsTask(method.ReturnType) || HasCustomAttribute(method, "ExcelDna.Registration.ExcelAsyncFunctionAttribute")) &&
                   HasExcelHandleReturn(method, excelHandleExternalTypeNames);
        }

        public static bool HasCustomAttribute(ISymbol symbol, string attribute)
        {
            return symbol.GetAttributes().Any(a => a.AttributeClass != null &&
                    Util.TypeHasAncestorWithFullName(a.AttributeClass, attribute));
        }

        public static bool HasReturnCustomAttribute(IMethodSymbol methodSymbol, string attribute)
        {
            return methodSymbol.GetReturnTypeAttributes().Any(a => a.AttributeClass != null &&
                    Util.TypeHasAncestorWithFullName(a.AttributeClass, attribute));
        }

        public static string CreateFunc16Args(IMethodSymbol method)
        {
            List<ITypeSymbol?> allParamTypes = method.Parameters.Take(method.Parameters.Length - 1).Select(p => p.Type).Cast<ITypeSymbol?>().ToList();
            var toAdd = 16 - allParamTypes.Count;
            for (int i = 0; i < toAdd; i++)
            {
                allParamTypes.Add(null);
            }
            allParamTypes.Add(method.ReturnType);

            return string.Join(",", allParamTypes.Select(i => i == null ? "object" : GetFullTypeName(i)));
        }

        static string BuildMethodType(IEnumerable<string> parameterTypeNames, ITypeSymbol returnType, bool returnsVoid)
        {
            return BuildMethodType("Action", "Func", parameterTypeNames, GetFullTypeName(returnType), returnsVoid);
        }

        static string BuildExtendedMethodType(IEnumerable<string> parameterTypeNames, ITypeSymbol returnType, bool returnsVoid)
        {
            return BuildExtendedMethodType(parameterTypeNames, GetFullTypeName(returnType), returnsVoid);
        }

        static string BuildMethodType(IEnumerable<string> parameterTypeNames, string returnTypeName, bool returnsVoid)
        {
            return BuildMethodType("Action", "Func", parameterTypeNames, returnTypeName, returnsVoid);
        }

        static string BuildExtendedMethodType(IEnumerable<string> parameterTypeNames, string returnTypeName, bool returnsVoid)
        {
            var parameterTypeNamesList = parameterTypeNames.ToList();
            return BuildMethodType($"ExcelDna.Integration.ExtendedAction{parameterTypeNamesList.Count}", $"ExcelDna.Integration.ExtendedFunc{parameterTypeNamesList.Count}", parameterTypeNamesList, returnTypeName, returnsVoid);
        }

        static string BuildMethodType(string action, string func, IEnumerable<string> parameterTypeNames, string returnTypeName, bool returnsVoid)
        {
            string parameters = string.Join(", ", parameterTypeNames);
            return returnsVoid
                ? $"{action}{(string.IsNullOrWhiteSpace(parameters) ? null : $"<{parameters}>")}"
                : $"{func}<{(string.IsNullOrWhiteSpace(parameters) ? null : $"{parameters}, ")}{returnTypeName}>";
        }

        static IEnumerable<string> GetAsyncWrapperParameterTypeNames(IMethodSymbol method, IEnumerable<IMethodSymbol>? userParameterConversions, IEnumerable<string>? excelHandleExternalTypeNames)
        {
            IEnumerable<IParameterSymbol> parameters = method.Parameters;
            if (parameters.Any() && IsCancellationToken(parameters.Last()))
                parameters = parameters.Take(parameters.Count() - 1);

            return parameters.Select(p => GetPostParameterConversionInputTypeName(p, userParameterConversions, excelHandleExternalTypeNames));
        }

        static string GetPostParameterConversionInputTypeName(IParameterSymbol parameter, IEnumerable<IMethodSymbol>? userParameterConversions, IEnumerable<string>? excelHandleExternalTypeNames)
        {
            // These conversions are part of the default registration processing pipeline and
            // change the lambda input shape seen by ParameterConversionRegistration.ApplyConversions.
            if (HasCustomAttribute(parameter, "ExcelDna.Integration.ExcelHandleAttribute") || TypeHasExcelHandle(parameter.Type, excelHandleExternalTypeNames))
                return "object";

            if (parameter.IsOptional || IsRangeType(parameter.Type))
                return "object";

            string parameterTypeName = GetFullTypeName(parameter.Type);
            if (userParameterConversions != null)
            {
                foreach (IMethodSymbol conversion in userParameterConversions.OrderBy(i => i.Name))
                {
                    if (conversion.Parameters.Length == 1 && parameterTypeName == GetFullTypeName(conversion.ReturnType))
                        parameterTypeName = GetFullTypeName(conversion.Parameters[0].Type);
                }
            }

            if (parameterTypeName != GetFullTypeName(parameter.Type))
                return parameterTypeName;

            if (IsStringArray(parameter.Type) || IsComplex(parameter.Type))
                return "object[]";

            if (IsString2DArray(parameter.Type))
                return "object[,]";

            if (parameter.Type.TypeKind == TypeKind.Enum || IsNullableType(parameter.Type))
                return "object";

            return GetFullTypeName(parameter.Type);
        }

        static string GetPostReturnConversionReturnTypeName(IMethodSymbol method, IEnumerable<IMethodSymbol>? userReturnConversions)
        {
            string returnTypeName = GetFullTypeName(method.ReturnType);
            if (userReturnConversions != null)
            {
                foreach (IMethodSymbol conversion in userReturnConversions.OrderBy(i => i.Name))
                {
                    if (conversion.Parameters.Length == 1 && returnTypeName == GetFullTypeName(conversion.Parameters[0].Type))
                        returnTypeName = GetFullTypeName(conversion.ReturnType);
                }
            }

            if (returnTypeName != GetFullTypeName(method.ReturnType))
                return returnTypeName;

            if (method.ReturnType.TypeKind == TypeKind.Enum)
                return "string";

            if (IsComplex(method.ReturnType))
                return "double[]";

            return returnTypeName;
        }

        static string GetAsyncObjectHandleAdapterReturnTypeName(IMethodSymbol method)
        {
            return IsTask(method.ReturnType) ? "System.Threading.Tasks.Task<string>" : "string";
        }

        static bool TypeHasExcelHandle(ITypeSymbol type, IEnumerable<string>? excelHandleExternalTypeNames)
        {
            return HasCustomAttribute(type, "ExcelDna.Integration.ExcelHandleAttribute") ||
                   (excelHandleExternalTypeNames != null && excelHandleExternalTypeNames.Contains(GetFullTypeName(type)));
        }

        static bool IsTask(ITypeSymbol type)
        {
            return type is INamedTypeSymbol namedType &&
                   namedType.IsGenericType &&
                   GetFullGenericTypeName(namedType) == "System.Threading.Tasks.Task";
        }

        static bool IsObservable(ITypeSymbol type)
        {
            return type is INamedTypeSymbol namedType &&
                   namedType.IsGenericType &&
                   GetFullGenericTypeName(namedType) == "System.IObservable";
        }

        static bool IsCancellationToken(IParameterSymbol parameter)
        {
            return GetFullTypeName(parameter.Type) == "System.Threading.CancellationToken";
        }

        static bool IsNullableType(ITypeSymbol type)
        {
            return type is INamedTypeSymbol namedType &&
                   namedType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T;
        }

        static bool IsRangeType(ITypeSymbol type)
        {
            return TypeHasAncestorWithFullName(type, "Microsoft.Office.Interop.Excel.Range") ||
                   TypeHasAncestorWithFullName(type, "ExcelDna.Integration.IRange");
        }

        static bool IsComplex(ITypeSymbol type)
        {
            return type.ToDisplayString(FullNameFormat) == "System.Numerics.Complex";
        }

        static bool IsStringArray(ITypeSymbol type)
        {
            return type is IArrayTypeSymbol arrayType &&
                   arrayType.Rank == 1 &&
                   arrayType.ElementType.SpecialType == SpecialType.System_String;
        }

        static bool IsString2DArray(ITypeSymbol type)
        {
            return type is IArrayTypeSymbol arrayType &&
                   arrayType.Rank == 2 &&
                   arrayType.ElementType.SpecialType == SpecialType.System_String;
        }

        private static SymbolDisplayFormat FullNameFormat = new SymbolDisplayFormat(typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces);
        private static SymbolDisplayFormat FullGenericNameFormat = new SymbolDisplayFormat(typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces, genericsOptions: SymbolDisplayGenericsOptions.None);
        private static SymbolDisplayFormat FullTypeOfTypeNameFormat = new SymbolDisplayFormat(
            typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
            genericsOptions: SymbolDisplayGenericsOptions.IncludeTypeParameters,
            miscellaneousOptions: SymbolDisplayMiscellaneousOptions.EscapeKeywordIdentifiers | SymbolDisplayMiscellaneousOptions.UseSpecialTypes);
    }
}
