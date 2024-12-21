﻿using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDna.AddIn.RegistrationSampleRuntimeTests
{
    public static class AsyncFunctionExamples
    {
        // Will not be registered in Excel by Excel-DNA, without being picked up by our Registration processing
        // since there is no ExcelFunction attribute, and ExplicitRegistration="true" in the .dna file prevents this 
        // function from being registered by the default processing.
        public static string dnaSayHello(string name)
        {
            return "Hello " + name + "!";
        }

        // Will be picked up by our explicit processing, no conversions applied, and normal registration
        [ExcelFunction(Name = "dnaSayHello")]
        public static string dnaSayHello2(string name)
        {
            if (name == "Bang!") throw new ArgumentException("Bad name!");
            return "Hello " + name + "!";
        }

        // A simple function that can take a long time to complete.
        // Will be wrapped to RunAsTask, via Task.Factory.StartNew(...)
        [ExcelAsyncFunction(Name = "dnaDelayedHello")]
        public static string dnaDelayedHello(string name, int msToSleep)
        {
            Thread.Sleep(msToSleep);
            return "Hello " + name + "!";
        }
    }
}