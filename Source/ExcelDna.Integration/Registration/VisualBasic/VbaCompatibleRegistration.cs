#if !COM_GENERATED

using ExcelDna.Integration;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelDna.Registration.VisualBasic
{
    // Performs function lookup and registration selected for VBA compatibility
    // * Does not require function to be marked by ExcelFunction
    // * Enables optional parameters and default values
    // * Enables Params parameter arrays
    // * Does ReferenceToRange conversions (including setting IsMacroType=True for such functions)
    [Microsoft.VisualBasic.CompilerServices.StandardModule]
    public sealed class VbaCompatibleRegistration
    {
        public static void PerformDefaultRegistration()
        {
            var conversionConfig = new ParameterConversionConfiguration()
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: false))
                .AddParameterConversion(RangeParameterConversion.ParameterConversion, null)
                ;

            GetAllPublicSharedFunctions()
                .ProcessParamsRegistrations()
                .UpdateRegistrationsForRangeParameters()
                .ProcessParameterConversions(conversionConfig)
                .RegisterFunctions();

            GetAllPublicSharedSubs().RegisterCommands();
        }

        // Gets the Public Shared methods that don't return Void
        private static IEnumerable<ExcelFunctionRegistration> GetAllPublicSharedFunctions()
        {
            return from ass in ExcelIntegration.GetExportedAssemblies()
                   from typ in ass.GetTypes()
                   where !typ.FullName.Contains(".My.") && typ.IsPublic
                   from mi in typ.GetMethods(BindingFlags.Public | BindingFlags.Static)
                   where mi.ReturnType != typeof(void) && !mi.IsSpecialName // Remove Property get_xxxx methods
                   select new ExcelFunctionRegistration(mi);
        }

        // Gets the Public Shared methods that return Void
        private static IEnumerable<ExcelCommandRegistration> GetAllPublicSharedSubs()
        {
            return from ass in ExcelIntegration.GetExportedAssemblies()
                   from typ in ass.GetTypes()
                   where !typ.FullName.Contains(".My.") && typ.IsPublic
                   from mi in typ.GetMethods(BindingFlags.Public | BindingFlags.Static)
                   where mi.ReturnType == typeof(void) && !mi.IsSpecialName
                   select new ExcelCommandRegistration(mi);
        }
    }
}

#endif
