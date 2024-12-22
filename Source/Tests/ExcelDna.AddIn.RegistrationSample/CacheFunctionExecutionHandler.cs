using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using ExcelDna.Registration;

namespace ExcelDna.AddIn.RegistrationSample
{
    // Instead of the Timeout as part of each Cache attribute, we could also have it as a global setting...
    [AttributeUsage(AttributeTargets.Method)]
    public class CacheAttribute : Attribute
    {
        public TimeSpan CacheTimeout;
        public CacheAttribute(int cacheTimeoutSeconds)
        {
            CacheTimeout = TimeSpan.FromSeconds(cacheTimeoutSeconds);
        }
    }

    // TODO: Implement simpler dictionary-based cache, using Excel-DNA hash keys from parameters.
    public class CacheFunctionExecutionHandler : FunctionExecutionHandler
    {
        static readonly ObjectCache _cache = MemoryCache.Default;
        readonly TimeSpan _cacheTimeout;

        public CacheFunctionExecutionHandler(TimeSpan cacheTimeout)
        {
            _cacheTimeout = cacheTimeout;
        }

        public override void OnEntry(FunctionExecutionArgs args)
        {
            string key = GenerateKey(args);
            args.Tag = key;

            string logMsg;
            object result = _cache[key];
            if (result != null)
            {
                // Set the return value and FlowBehavior, to short-cut the function call
                args.ReturnValue = result;
                args.FlowBehavior = FlowBehavior.Return;
                logMsg = $"CacheFunctionExecutionHandler {args.FunctionName} result in cache {key}.";
            }
            else
            {
                logMsg = $"CacheFunctionExecutionHandler {args.FunctionName} result not in cache {key}.";
            }

            Debug.Print(logMsg);
            Logger.Log(logMsg);
        }

        public override void OnSuccess(FunctionExecutionArgs args)
        {
            // Store in cache
            _cache.Add((string)args.Tag, args.ReturnValue,
                new CacheItemPolicy { AbsoluteExpiration = DateTime.UtcNow + _cacheTimeout });
        }

        string GenerateKey(FunctionExecutionArgs args)
        {
            var keyBuilder = new StringBuilder(args.FunctionName);
            foreach (var arg in args.Arguments)
            {
                keyBuilder.Append('\0');    // 0-delimited !?
                keyBuilder.Append(arg);
            }
            return keyBuilder.ToString();
        }

        /////////////////////// Registration handler //////////////////////////////////
        // (This code can be anywhere... - need not be in this class)

        // In this case, we only ever make one 'handler' object
        public static FunctionExecutionHandler CacheHandlerSelector(ExcelFunctionRegistration functionRegistration)
        {
            // Eat the TimingAttributes, and return a timer handler if there were any
            if (functionRegistration.CustomAttributes.OfType<CacheAttribute>().Any())
            {
                // Get the first cache attribute, and remove all of them
                var cacheAtt = functionRegistration.CustomAttributes.OfType<CacheAttribute>().First();
                functionRegistration.CustomAttributes.RemoveAll(att => att is CacheAttribute);

                return new CacheFunctionExecutionHandler(cacheAtt.CacheTimeout);
            }
            return null;
        }

    }


}

