using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelDna.Registration
{
    public class ExcelParameterRegistration : IExcelFunctionParameter
    {
        // Used for the final Excel-DNA registration
        public ExcelArgumentAttribute ArgumentAttribute { get; private set; }

        // Used only for the Registration processing
        public List<object> CustomAttributes { get; private set; } // Should not be null, and elements should not be null

        public ExcelParameterRegistration(ExcelArgumentAttribute argumentAttribute)
        {
            if (argumentAttribute == null) throw new ArgumentNullException("argumentAttribute");
            ArgumentAttribute = argumentAttribute;

            CustomAttributes = new List<object>();
        }

        /// <summary>
        /// Also craetes attributes from Optional / Default Value
        /// </summary>
        /// <param name="parameterInfo"></param>
        public ExcelParameterRegistration(ParameterInfo parameterInfo)
        {
            CustomAttributes = new List<object>();

            var allParameterAttributes = parameterInfo.GetCustomAttributes(true);
            foreach (var att in allParameterAttributes)
            {
                var argAtt = att as ExcelArgumentAttribute;
                if (argAtt != null)
                {
                    ArgumentAttribute = argAtt;
                    if (string.IsNullOrEmpty(ArgumentAttribute.Name))
                        ArgumentAttribute.Name = parameterInfo.Name;
                }
                else
                {
                    CustomAttributes.Add(att);
                }
            }

            CustomAttributes.AddRange(ExcelTypeDescriptor.GetCustomAttributes(parameterInfo.ParameterType));

            // Check that the ExcelArgumentAttribute has been set
            if (ArgumentAttribute == null)
            {
                ArgumentAttribute = new ExcelArgumentAttribute { Name = parameterInfo.Name };
            }

            object parameterDefaultValue = ParameterDefaultValue(parameterInfo);
            // Extra processing for Optional / Default values
            // TODO: Also consider DefaultValueAttribute (which is wrong, but might be used...)
            if (parameterInfo.IsOptional && parameterDefaultValue != DBNull.Value)
            {
                Debug.Assert(CustomAttributes.OfType<OptionalAttribute>().Any());
                Debug.Assert(!CustomAttributes.OfType<DefaultParameterValueAttribute>().Any());
                CustomAttributes.Add(new DefaultParameterValueAttribute(parameterDefaultValue));
            }
        }

        // Checks that the property invariants are met, particularly regarding the attributes lists.
        internal bool IsValid()
        {
            return ArgumentAttribute != null && CustomAttributes != null && CustomAttributes.All(att => att != null);
        }

        private static object ParameterDefaultValue(ParameterInfo parameterInfo)
        {
#if NETFRAMEWORK
            // A workaround for optional DateTime parameter's default value exception https://github.com/dotnet/runtime/issues/24574
            if (parameterInfo.ParameterType == typeof(DateTime) && parameterInfo.Attributes == (ParameterAttributes.Optional | ParameterAttributes.HasDefault))
            {
                try
                {
                    return parameterInfo.DefaultValue;
                }
                catch (FormatException)
                {
                    return null;
                }
            }
#endif

            return parameterInfo.DefaultValue;
        }
    }
}
