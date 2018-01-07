//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;

namespace ExcelDna.Integration
{
	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
	public class ExcelFunctionAttribute : Attribute
	{
		private string _description;

		public virtual string Name { get; set; }

		public virtual string Description
		{
			get { return _description; }
			set { _description = value; }
		}

		public virtual string Category { get; set; }
		public virtual string HelpTopic { get; set; }
		public virtual bool IsVolatile { get; set; }
		public virtual bool IsHidden { get; set; }
		public virtual bool IsExceptionSafe { get; set; }
		public virtual bool IsMacroType { get; set; }
		public virtual bool IsThreadSafe { get; set; }
		public virtual bool IsClusterSafe { get; set; }
		public virtual bool ExplicitRegistration { get; set; }
		public virtual bool SuppressOverwriteError { get; set; }

		public ExcelFunctionAttribute()
		{
		}

		public ExcelFunctionAttribute(string description)
		{
			_description = description;
		}
	}

	[AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	public class ExcelArgumentAttribute : Attribute
	{
		private string _description;

		public virtual string Name { get; set; }

		public virtual string Description
		{
			get { return _description; }
			set { _description = value; }
		}

		public virtual bool AllowReference { get; set; }

		public ExcelArgumentAttribute()
		{
		}

		public ExcelArgumentAttribute(string description)
		{
			_description = description;
		}
	}

	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
	public class ExcelCommandAttribute : Attribute
	{
		private string _description;

		public virtual string Name { get; set; }

		public virtual string Description
		{
			get { return _description; }
			set { _description = value; }
		}

		public virtual string HelpTopic { get; set; }
		public virtual string ShortCut { get; set; }
		public virtual string MenuName { get; set; }
		public virtual string MenuText { get; set; }
		public virtual bool IsExceptionSafe { get; set; }
		public virtual bool ExplicitRegistration { get; set; }
		public virtual bool SuppressOverwriteError { get; set; }

		[Obsolete("ExcelFunctions can be declared hidden, not ExcelCommands.")]
		public bool IsHidden = false;

		public ExcelCommandAttribute()
		{
		}

		public ExcelCommandAttribute(string description)
		{
			_description = description;
		}
	}
}
