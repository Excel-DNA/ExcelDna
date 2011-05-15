/*
  Copyright (C) 2005-2011 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

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
