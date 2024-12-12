using System;
using System.Collections.Generic;

namespace ExcelDna.Integration
{
    internal class ExcelTypeDescriptor
    {
        private static Dictionary<Type, List<object>> typeAttributes = new Dictionary<Type, List<object>>();

        public static void AddCustomAttributes(Type t, IEnumerable<object> attributes)
        {
            if (!typeAttributes.ContainsKey(t))
                typeAttributes.Add(t, new List<object>());

            typeAttributes[t].AddRange(attributes);
        }

        public static List<object> GetCustomAttributes(Type t)
        {
            List<object> result = new List<object>();
            result.AddRange(t.GetCustomAttributes(true));

            if (typeAttributes.ContainsKey(t))
                result.AddRange(typeAttributes[t]);

            return result;
        }
    }
}
