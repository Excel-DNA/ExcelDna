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
            string parameters = string.Join(", ", method.Parameters.Select(p => GetFullTypeName(p.Type)));
            return method.ReturnsVoid ?
                $"Action{(string.IsNullOrWhiteSpace(parameters) ? null : $"<{parameters}>")}" :
                $"Func<{(string.IsNullOrWhiteSpace(parameters) ? null : $"{parameters}, ")}{GetFullTypeName(method.ReturnType)}>";
        }

        public static bool IsLastArrayParams(IMethodSymbol method)
        {
            return method.Parameters.Length > 0 && method.Parameters.Last().IsParams && method.Parameters.Last().Type is IArrayTypeSymbol;
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

        private static SymbolDisplayFormat FullNameFormat = new SymbolDisplayFormat(typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces);
    }
}
