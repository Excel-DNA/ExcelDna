using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelDna.Registration
{
    // CONSIDER: Improve safety here... make invalid data unrepresentable.
    // CONSIDER: Should ExcelCommands also be handled here...? For the moment not...
    public class ExcelFunctionRegistration : IExcelFunctionInfo
    {
        // These are used for registration
        public LambdaExpression FunctionLambda { get; set; }
        public ExcelFunctionAttribute FunctionAttribute { get; set; }                   // May not be null
        public List<ExcelParameterRegistration> ParameterRegistrations { get; private set; }    // A list of ExcelParameterRegistrations with length equal to the number of parameters in Delegate

        // These are used only for the Registration processing
        public List<object> CustomAttributes { get; private set; }                 // List may not be null
        public ExcelReturnRegistration ReturnRegistration { get; private set; }

        public List<ExcelParameterRegistration> Parameters => ParameterRegistrations.Cast<ExcelParameterRegistration>().ToList();
        public IExcelFunctionReturn Return => ReturnRegistration;

        // Checks that the property invariants are met, particularly regarding the attributes lists.
        internal bool IsValid()
        {
            return FunctionLambda != null &&
                   FunctionAttribute != null &&
                   ParameterRegistrations != null &&
                   ParameterRegistrations.Count == FunctionLambda.Parameters.Count &&
                   CustomAttributes != null &&
                   CustomAttributes.All(att => att != null) &&
                   ReturnRegistration != null &&
                   ReturnRegistration.CustomAttributes != null &&
                   ReturnRegistration.CustomAttributes.All(att => att != null) &&
                   ParameterRegistrations.All(pr => pr.IsValid());
        }

        /// <summary>
        /// Creates a new ExcelFunctionRegistration with the given LambdaExpression.
        /// Uses the passes in attributes for registration.
        /// 
        /// The number of ExcelParameterRegistrations passed in must match the number of parameters in the LambdaExpression.
        /// </summary>
        /// <param name="functionLambda"></param>
        /// <param name="functionAttribute"></param>
        /// <param name="parameterRegistrations"></param>
        public ExcelFunctionRegistration(LambdaExpression functionLambda, ExcelFunctionAttribute functionAttribute, IEnumerable<ExcelParameterRegistration> parameterRegistrations = null)
        {
            if (functionLambda == null) throw new ArgumentNullException("functionLambda");
            if (functionAttribute == null) throw new ArgumentNullException("functionAttribute");

            FunctionLambda = functionLambda;
            FunctionAttribute = functionAttribute;
            if (parameterRegistrations == null)
            {
                if (functionLambda.Parameters.Count != 0) throw new ArgumentOutOfRangeException("parameterRegistrations", "No parameter registrations provided, but function has parameters.");
                ParameterRegistrations = new List<ExcelParameterRegistration>();
            }
            else
            {
                ParameterRegistrations = new List<ExcelParameterRegistration>(parameterRegistrations);
                if (functionLambda.Parameters.Count != ParameterRegistrations.Count) throw new ArgumentOutOfRangeException("parameterRegistrations", "Mismatched number of ParameterRegistrations provided.");
            }

            // Create the lists - hope the rest is filled in right...?
            CustomAttributes = new List<object>();
            ReturnRegistration = new ExcelReturnRegistration();
        }

        /// <summary>
        /// Creates a new ExcelFunctionRegistration from a LambdaExpression.
        /// Uses the Name and Parameter Names to fill in the default attributes.
        /// </summary>
        /// <param name="functionLambda"></param>
        public ExcelFunctionRegistration(LambdaExpression functionLambda)
        {
            if (functionLambda == null) throw new ArgumentNullException("functionLambda");

            FunctionLambda = functionLambda;
            FunctionAttribute = new ExcelFunctionAttribute { Name = functionLambda.Name };
            ParameterRegistrations = functionLambda.Parameters
                                     .Select(p => new ExcelParameterRegistration(new ExcelArgumentAttribute { Name = p.Name }))
                                     .ToList();

            CustomAttributes = new List<object>();
            ReturnRegistration = new ExcelReturnRegistration();
        }

        // NOTE: 16 parameter max for Expression.GetDelegateType
        // Copies all the (non Excel...) attributes from the method into the CustomAttribute lists.
        // TODO: What about native async function, which returns 'Void'?

        /// <summary>
        /// Creates a new ExcelFunctionRegistration from a MethodInfo, with a LambdaExpression that represents a call to the method.
        /// Uses the Name and Parameter Names from the MethodInfo to fill in the default attributes.
        /// All CustomAttributes on the method and parameters are copies to the respective collections in the ExcelFunctionRegistration.
        /// </summary>
        /// <param name="methodInfo"></param>
        public ExcelFunctionRegistration(MethodInfo methodInfo)
        {
            CustomAttributes = new List<object>();

            var paramExprs = methodInfo.GetParameters()
                             .Select(pi => Expression.Parameter(pi.ParameterType, pi.Name))
                             .ToList();
            FunctionLambda = (paramExprs.Count > 16) ?
                Expression.Lambda(GetExtendedDelegateType(methodInfo), Expression.Call(methodInfo, paramExprs), methodInfo.Name, paramExprs) :
                Expression.Lambda(Expression.Call(methodInfo, paramExprs), methodInfo.Name, paramExprs);

            var allMethodAttributes = methodInfo.GetCustomAttributes(true);
            foreach (var att in allMethodAttributes)
            {
                var funcAtt = att as ExcelFunctionAttribute;
                if (funcAtt != null)
                {
                    FunctionAttribute = funcAtt;
                    // At least ensure that name is set - from the method if need be.
                    if (string.IsNullOrEmpty(FunctionAttribute.Name))
                        FunctionAttribute.Name = methodInfo.Name;
                }
                else
                {
                    CustomAttributes.Add(att);
                }
            }
            // Check that ExcelFunctionAttribute has been set
            if (FunctionAttribute == null)
            {
                FunctionAttribute = new ExcelFunctionAttribute { Name = methodInfo.Name };
            }

            ParameterRegistrations = methodInfo.GetParameters().Select(pi => new ExcelParameterRegistration(pi)).ToList();
            ReturnRegistration = new ExcelReturnRegistration();
            ReturnRegistration.CustomAttributes.AddRange(methodInfo.ReturnParameter.GetCustomAttributes(true));

            Type returnType = methodInfo.ReturnType;
            if (returnType.IsGenericType && (returnType.GetGenericTypeDefinition() == typeof(Task<>) || returnType.GetGenericTypeDefinition() == typeof(IObservable<>)))
                returnType = returnType.GetGenericArguments()[0];
            ReturnRegistration.CustomAttributes.AddRange(ExcelTypeDescriptor.GetCustomAttributes(returnType));

            // Check that we haven't made a mistake
            Debug.Assert(IsValid());
        }

        private static Type GetExtendedDelegateType(MethodInfo methodInfo)
        {
            if (methodInfo.ReturnType != typeof(void))
            {
                Type genericBase = ExtendedFuncUtil.GetFuncType(methodInfo.GetParameters().Length);
                var args = methodInfo.GetParameters().Select(p => p.ParameterType)
                    .Concat(new[] { methodInfo.ReturnType })
                    .ToArray();
                return genericBase.MakeGenericType(args);
            }
            else
            {
                Type genericBase = ExtendedFuncUtil.GetActionType(methodInfo.GetParameters().Length);
                if (!genericBase.IsGenericType)
                    return genericBase;

                var args = methodInfo.GetParameters().Select(p => p.ParameterType)
                    .ToArray();
                return genericBase.MakeGenericType(args);
            }
        }
    }
}
