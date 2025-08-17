using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelDna.Integration;

namespace ExcelDna.Registration
{
    // Explicit support here for ExcelCommands is to encourage ExplicitRegistration=true
    // for all add-ins that use the Registration processing.
    // But to support this we need to take care of ExcelCommands explicitly too.

    // Maybe one day we'll do Command/Function unification
    // For now we mirror core Excel-DNA approach
    // Note that Excel-DNA does support ExcelCommands that take parameters and return values.
    // However, these are not available as worksheet functions, and are unusual - 
    // so the ExcelCommandRegistration here doesn't support attributes on such parameters or return values.
    public class ExcelCommandRegistration
    {
        // These are used for registration
        public LambdaExpression CommandLambda { get; set; }
        public ExcelCommandAttribute CommandAttribute { get; set; }        // May not be null

        // These are used only for the Registration processing
        public List<object> CustomAttributes { get; set; }                 // List may not be null

        public ExcelCommandRegistration(LambdaExpression commandLambda, ExcelCommandAttribute commandAttribute)
        {
            if (commandLambda == null) throw new ArgumentNullException("commandLambda");
            if (commandAttribute == null) throw new ArgumentNullException("commandAttribute");

            CommandLambda = commandLambda;
            CommandAttribute = commandAttribute;

            // Create the lists - hope the rest is filled in right...?
            CustomAttributes = new List<object>();
        }

        public ExcelCommandRegistration(LambdaExpression commandLambda)
        {
            if (commandLambda == null) throw new ArgumentNullException("commandLambda");

            CommandLambda = commandLambda;
            CommandAttribute = new ExcelCommandAttribute { Name = commandLambda.Name };
            CustomAttributes = new List<object>();
        }

        public ExcelCommandRegistration(MethodInfo methodInfo)
        {
            CustomAttributes = new List<object>();

            var paramExprs = methodInfo.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();
            CommandLambda = Expression.Lambda(Expression.Call(methodInfo, paramExprs), methodInfo.Name, paramExprs);

            var allMethodAttributes = methodInfo.GetCustomAttributes(true);
            foreach (var att in allMethodAttributes)
            {
                var cmdAtt = att as ExcelCommandAttribute;
                if (cmdAtt != null)
                {
                    CommandAttribute = cmdAtt;
                    // At least ensure that name is set - from the method if need be.
                    if (string.IsNullOrEmpty(CommandAttribute.Name))
                        CommandAttribute.Name = methodInfo.Name;
                }
                else
                {
                    CustomAttributes.Add(att);
                }
            }
            // Check that ExcelCommandAttribute has been set
            if (CommandAttribute == null)
            {
                CommandAttribute = new ExcelCommandAttribute { Name = methodInfo.Name };
            }
        }

        internal static bool IsCommand(MethodInfo methodInfo)
        {
            return methodInfo.GetCustomAttribute<ExcelCommandAttribute>() != null;
        }
    }
}
