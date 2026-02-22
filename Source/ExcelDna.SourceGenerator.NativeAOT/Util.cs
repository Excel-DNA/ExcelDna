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

        public static string MethodExpressionType(IMethodSymbol method)
        {
            return $"System.Linq.Expressions.Expression<{MethodType(method)}>";
        }

        public static string MethodPostParameterConversionType(IMethodSymbol method)
        {
            return BuildMethodType(method.Parameters.Select(GetPostParameterConversionInputTypeName), method.ReturnType, method.ReturnsVoid);
        }

        public static bool HasPostParameterConversionShape(IMethodSymbol method)
        {
            return MethodPostParameterConversionType(method) != MethodType(method);
        }

        public static bool IsLastArrayParams(IMethodSymbol method)
        {
            return method.Parameters.Length > 0 && method.Parameters.Last().IsParams && method.Parameters.Last().Type is IArrayTypeSymbol;
        }


        public static bool HasCustomAttribute(IMethodSymbol methodSymbol, string attribute)
        {
            return methodSymbol.GetAttributes().Any(a => a.AttributeClass != null &&
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
            string parameters = string.Join(", ", parameterTypeNames);
            return returnsVoid
                ? $"Action{(string.IsNullOrWhiteSpace(parameters) ? null : $"<{parameters}>")}"
                : $"Func<{(string.IsNullOrWhiteSpace(parameters) ? null : $"{parameters}, ")}{GetFullTypeName(returnType)}>";
        }

        static string GetPostParameterConversionInputTypeName(IParameterSymbol parameter)
        {
            // These conversions are part of the default registration processing pipeline and
            // change the lambda input shape seen by ParameterConversionRegistration.ApplyConversions.
            if (parameter.IsOptional || parameter.Type.TypeKind == TypeKind.Enum || IsNullableType(parameter.Type) || IsRangeType(parameter.Type))
                return "object";

            if (IsStringArray(parameter.Type) || IsComplex(parameter.Type))
                return "object[]";

            if (IsString2DArray(parameter.Type))
                return "object[,]";

            return GetFullTypeName(parameter.Type);
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
    }
}
