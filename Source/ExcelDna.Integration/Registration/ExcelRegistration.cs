using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.Registration
{
    public static class ExcelRegistration
    {
        /// <summary>
        /// Retrieve registration wrappers for all (public, static) functions marked with [ExcelFunction] attributes, 
        /// in all exported assemblies.
        /// </summary>
        /// <returns>All public static methods in registered assemblies that are decorated with an [ExcelFunction] attribute 
        /// (or a derived attribute, like [ExcelAsyncFunction]).
        /// </returns>
        public static IEnumerable<ExcelFunctionRegistration> GetExcelFunctions()
        {
            return from ass in ExcelIntegration.GetExportedAssemblies()
                   from typ in ass.GetTypes()
                   from mi in typ.GetMethods(BindingFlags.Public | BindingFlags.Static)
                   where mi.GetCustomAttribute<ExcelFunctionAttribute>() != null
                   select new ExcelFunctionRegistration(mi);
        }

        /// <summary>
        /// Registers the given functions with Excel-DNA.
        /// </summary>
        /// <param name="registrationEntries"></param>
        public static void RegisterFunctions(this IEnumerable<ExcelFunctionRegistration> registrationEntries)
        {
            ExcelDna.Integration.ExtendedRegistration.Registration.Register(registrationEntries);
        }

        /// <summary>
        /// Retrieve registration wrappers for all (public, static) methods marked with [ExcelCommand] attributes, 
        /// in all exported assemblies.
        /// </summary>
        /// <returns>All public static methods in registered assemblies that are decorated with an [ExcelCommand] attribute 
        /// (or a derived attribute)
        /// </returns>
        public static IEnumerable<ExcelCommandRegistration> GetExcelCommands()
        {
            return from ass in ExcelIntegration.GetExportedAssemblies()
                   from typ in ass.GetTypes()
                   from mi in typ.GetMethods(BindingFlags.Public | BindingFlags.Static)
                   where mi.GetCustomAttribute<ExcelCommandAttribute>() != null
                   select new ExcelCommandRegistration(mi);
        }

        /// <summary>
        /// Registers the given macros with Excel-DNA.
        /// </summary>
        /// <param name="registrationEntries"></param>
        public static void RegisterCommands(this IEnumerable<ExcelCommandRegistration> registrationEntries)
        {
            var lambdas = registrationEntries.Select(reg => reg.CommandLambda).ToList();
            var attribs = registrationEntries.Select(reg => reg.CommandAttribute).ToList<object>();

            ExcelIntegration.RegisterLambdaExpressions(lambdas, attribs, null);
        }
    }
}
