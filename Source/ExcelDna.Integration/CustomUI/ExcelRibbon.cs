﻿//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using ExcelDna.Logging;
using JetBrains.Annotations;
using CLSID = System.Guid;

namespace ExcelDna.Integration.CustomUI
{
    [ComVisible(true)]
#pragma warning disable CS0618 // Type or member is obsolete (but probably not forever)
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning restore CS0618 // Type or member is obsolete
    public class ExcelRibbon : ExcelComAddIn, IRibbonExtensibility //, ICustomTaskPaneConsumer
    {
        public const string NamespaceCustomUI2010 = @"http://schemas.microsoft.com/office/2009/07/customui";
        public const string NamespaceCustomUI2007 = @"http://schemas.microsoft.com/office/2006/01/customui";
        // Info about changes in Excel 2013: http://www.rondebruin.nl/win/s2/win005.htm

        [CanBeNull]
        public virtual string GetCustomUI(string RibbonID)
        {
            if (RibbonID != "Microsoft.Excel.Workbook")
            {
                Logger.ComAddIn.Error("ExcelRibbon.GetCustomUI - Invalid RibbonID for Excel. RibbonID: {0}", RibbonID);
                return null;
            }

            // Default behaviour for GetCustomUI is to look in DnaLibrary.
            // We return a CustomUI based on the version of Excel, as follows:
            // If Excel12, look for a CustomUI with namespace... If not found, return nothing
            // If Excel14 or bigger look for CustomUI with namespace ... If not found, look for ... else return nothing.
            // (not sure how to future-proof...)

            Dictionary<string, string> customUIs = new Dictionary<string, string>();
            foreach (XmlNode customUI in this.DnaLibrary.CustomUIs)
            {
                customUIs[customUI.NamespaceURI] = customUI.OuterXml;
            }

            if (ExcelDnaUtil.ExcelVersion >= 14.0)
            {
                if (customUIs.ContainsKey(NamespaceCustomUI2010))
                {
                    return customUIs[NamespaceCustomUI2010];
                }
                if (customUIs.ContainsKey(NamespaceCustomUI2007))
                {
                    return customUIs[NamespaceCustomUI2007];
                }
                return null;
            }
            if (ExcelDnaUtil.ExcelVersion >= 12.0)
            {
                if (customUIs.ContainsKey(NamespaceCustomUI2007))
                {
                    return customUIs[NamespaceCustomUI2007];
                }
                return null;
            }
            throw new InvalidOperationException("Not expected to provide CustomUI string for Excel version < 12.0");
        }

        // LoadImage helper - to use need to mark loadImage='LoadImage' in the xml.
        // 1. An IPictureDisp
        // 2. A System.Drawing.Bitmap
        // 3. A string containing an imageMso identifier
        // Our default implementation ...
        public virtual object LoadImage(string imageId)
        {
            // Default implementation ...
            return DnaLibrary.GetImage(imageId);
        }

        // RunTagMacro helper function
        public virtual void RunTagMacro(IRibbonControl control)
        {
            if (!string.IsNullOrEmpty(control.Tag))
            {
                // CONSIDER: Is this a danger for shutting down - surely not...?
                object app = ExcelDnaUtil.Application;
                app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, new object[] { control.Tag }, new CultureInfo(1033));
            }
        }
    }
}
