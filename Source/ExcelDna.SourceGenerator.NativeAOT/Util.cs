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

        private static SymbolDisplayFormat FullNameFormat = new SymbolDisplayFormat(typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces);
    }
}
