//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

namespace ExcelDna.Integration
{
	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
	public class ExcelFunctionAttribute : Attribute
	{
		public string Name = null;
		public string Description = null;
		public string Category = null;
		public string HelpTopic = null;
		public bool   IsVolatile = false;
        public bool   IsHidden = false;
		public bool   IsExceptionSafe = false;
		public bool   IsMacroType = false;
        public bool   IsThreadSafe = false;
        public bool   IsClusterSafe = false;
        public bool   ExplicitRegistration = false;

		public ExcelFunctionAttribute()
		{
		}

		public ExcelFunctionAttribute(string description)
		{
			Description = description;
		}
	}

	[AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	public class ExcelArgumentAttribute : Attribute
	{
		public string Name = null;
		public string Description = null;
		public bool   AllowReference = false;

		public ExcelArgumentAttribute()
		{
		}

		public ExcelArgumentAttribute(string description)
		{
			Description = description;
		}
	}

	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
	public class ExcelCommandAttribute : Attribute
	{
		public string Name = null;
		public string Description = null;
		public string HelpTopic = null;
		public string ShortCut = null;
		public string MenuName = null;
		public string MenuText = null;
        public bool IsExceptionSafe = false;
        public bool ExplicitRegistration = false;

        [Obsolete("ExcelFunctions can be declared hidden, not ExcelCommands.")]
		public bool IsHidden = false;

		public ExcelCommandAttribute()
		{
		}

		public ExcelCommandAttribute(string description)
		{
			Description = description;
		}
	}
}
