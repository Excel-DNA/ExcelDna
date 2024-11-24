using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal class ExcelParameter : IExcelFunctionParameter
    {
        // Used for the final Excel-DNA registration
        public ExcelArgumentAttribute ArgumentAttribute { get; private set; }

        // Used only for the Registration processing
        public List<object> CustomAttributes { get; private set; } // Should not be null, and elements should not be null

        public ExcelParameter(ExcelArgumentAttribute argumentAttribute)
        {
            if (argumentAttribute == null) throw new ArgumentNullException("argumentAttribute");
            ArgumentAttribute = argumentAttribute;

            CustomAttributes = new List<object>();
        }

        /// <summary>
        /// Also craetes attributes from Optional / Default Value
        /// </summary>
        /// <param name="parameterInfo"></param>
        public ExcelParameter(ParameterInfo parameterInfo)
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

            CustomAttributes.AddRange(parameterInfo.ParameterType.GetCustomAttributes(true));

            // Check that the ExcelArgumentAttribute has been set
            if (ArgumentAttribute == null)
            {
                ArgumentAttribute = new ExcelArgumentAttribute { Name = parameterInfo.Name };
            }

            // Extra processing for Optional / Default values
            // TODO: Also consider DefaultValueAttribute (which is wrong, but might be used...)
            if (parameterInfo.IsOptional && parameterInfo.DefaultValue != DBNull.Value)
            {
                Debug.Assert(CustomAttributes.OfType<OptionalAttribute>().Any());
                Debug.Assert(!CustomAttributes.OfType<DefaultParameterValueAttribute>().Any());
                CustomAttributes.Add(new DefaultParameterValueAttribute(parameterInfo.DefaultValue));
            }
        }

        // Checks that the property invariants are met, particularly regarding the attributes lists.
        internal bool IsValid()
        {
            return ArgumentAttribute != null && CustomAttributes != null && CustomAttributes.All(att => att != null);
        }
    }
}
