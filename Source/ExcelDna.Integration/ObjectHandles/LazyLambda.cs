using System;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelDna.Integration.ObjectHandles
{
    internal class LazyLambda
    {
        private LambdaExpression exp;
        private object[] args;

#if AOT_COMPATIBLE
        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL3050:RequiresDynamicCode", Justification = "Passes all tests")]
#endif
        public static LambdaExpression Create(LambdaExpression source)
        {
            var wrappingParameters = source.Parameters.Select(p => Expression.Parameter(p.Type, p.Name)).ToList();
            var lazyLambdaConstructor = typeof(LazyLambda).GetConstructor(new Type[] { typeof(LambdaExpression), typeof(object[]) });
            var parametersArray = Expression.NewArrayInit(typeof(object), wrappingParameters.Select(p => Expression.Convert(p, typeof(object))));
            var wrappingCall = Expression.New(lazyLambdaConstructor, new Expression[] { Expression.Constant(source), parametersArray });
            return Expression.Lambda(wrappingCall, wrappingParameters);
        }

        public LazyLambda(LambdaExpression exp, params object[] args)
        {
            this.exp = exp;
            this.args = args;
        }

        public object Invoke()
        {
            return exp.Compile().DynamicInvoke(args);
        }
    }
}
