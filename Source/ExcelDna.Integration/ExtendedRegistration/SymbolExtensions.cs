// Sweet helper to get MethodInfos without magic strings, 
// from Samual Jack here: http://blog.functionalfun.net/2009/10/getting-methodinfo-of-generic-method.html

using System.Linq.Expressions;
using System.Reflection;
using System;

namespace ExcelDna.Integration.ExtendedRegistration
{
    internal static class SymbolExtensions
    {
        /// <summary>
        /// Given a lambda expression that calls a method, returns the method info.
        /// </summary>
        /// <param name="expression">The expression.</param>
        /// <returns></returns>
        public static MethodInfo GetMethodInfo(Expression<Action> expression)
        {
            return GetMethodInfo((LambdaExpression)expression);
        }

        /// <summary>
        /// Given a lambda expression that calls a method, returns the method info.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="expression">The expression.</param>
        /// <returns></returns>
        public static MethodInfo GetMethodInfo<T>(Expression<Action<T>> expression)
        {
            return GetMethodInfo((LambdaExpression)expression);
        }

        /// <summary>
        /// Given a lambda expression that calls a method, returns the method info.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="expression">The expression.</param>
        /// <returns></returns>
        public static MethodInfo GetMethodInfo<T, TResult>(Expression<Func<T, TResult>> expression)
        {
            return GetMethodInfo((LambdaExpression)expression);
        }

        /// <summary>
        /// Given a lambda expression that calls a method, returns the method info.
        /// </summary>
        /// <param name="expression">The expression.</param>
        /// <returns></returns>
        public static MethodInfo GetMethodInfo(LambdaExpression expression)
        {
            MethodCallExpression outermostExpression = expression.Body as MethodCallExpression;

            if (outermostExpression == null)
            {
                throw new ArgumentException("Invalid Expression. Expression should consist of a Method call only.");
            }

            return outermostExpression.Method;
        }

        public static MemberExpression GetProperty<T, TValue>(ParameterExpression instance, Expression<Func<T, TValue>> propertyLambda)
        {
            var propertyMember = propertyLambda.Body as MemberExpression;
            if (propertyMember == null)
                throw new ArgumentException("Invalid Expression. Expression should consist of a single member call only.");

            var propertyInfo = propertyMember.Member as PropertyInfo;
            if (propertyInfo == null)
                throw new ArgumentException("Invalid Expression. Expression should consist of a Property call only.");

            if (!propertyInfo.ReflectedType.IsAssignableFrom(instance.Type))
                throw new ArgumentException(string.Format(
                    "Expresion '{0}' refers to a property that is not from type {1}.", propertyLambda, instance.Type));

            return Expression.Property(instance, propertyInfo);
        }
    }
}
