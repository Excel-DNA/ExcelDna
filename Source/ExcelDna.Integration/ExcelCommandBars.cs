//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using ExcelDna.Serialization;

namespace ExcelDna.Integration.CustomUI
{
    public delegate Bitmap GetImageDelegate(string imageName);

    //public class ExcelCommandBars
    //{
    //    public ExcelCommandBars()
    //    {
    //    }

    //    public virtual string GetCustomUI()
    //    {
    //    }

    //    public virtual object GetImage(string imageName)
    //    {

    //    }

    //    public CommandBarControls Controls
    //    {
    //        get
    //        {
    //        }
    //    }
    //}

    public static class ExcelCommandBarUtil
    {
        // List of loaded CustomUI 
        static List<XmlNode> loadedCustomUIs = new List<XmlNode>();

        // Helper to call Application.CommandBars
        public static CommandBars GetCommandBars()
        {
            Application excelApp = new Application(ExcelDnaUtil.Application);
            return excelApp.CommandBars;
        }

        public static void LoadCommandBars(string xmlCustomUI)
        {
            LoadCommandBars(xmlCustomUI, delegate(string imageName) { return null; });
        }

        public static void LoadCommandBars(string xmlCustomUI, GetImageDelegate getImage)
        {
            string dnaLibraryWrapper = string.Format(@"<DnaLibrary><CustomUI>{0}</CustomUI></DnaLibrary>", xmlCustomUI);
            using (StringReader sr = new StringReader(dnaLibraryWrapper))
            {
                XmlSerializer serializer = new DnaLibrarySerializer();
                // TODO: Check for and display errors....
                DnaLibrary dnaLibrary = (DnaLibrary)serializer.Deserialize(sr);
                LoadCommandBars(dnaLibrary.CustomUIs[0], getImage);
            }
        }

        internal static void LoadCommandBars(XmlNode xmlCustomUI, GetImageDelegate getImage)
        {
            if (xmlCustomUI.NamespaceURI != "http://schemas.excel-dna.net/office/2003/01/commandbars")
            {
                // Unsupported version ....
                // TODO: Log display ....?
                Debug.Print("Unsupported commandBars version.");
                return;
            }

            Application excelApp = new Application(ExcelDnaUtil.Application);

            loadedCustomUIs.Add(xmlCustomUI);
            try
            {
                AddCommandBarControls(excelApp, xmlCustomUI.ChildNodes, getImage);
            }
            catch (Exception e)
            {
                // Suppress exceptions
                Debug.Print("ExcelCommandBars: Error adding controls: {0}", e);
            }
        }

        public static void UnloadCommandBars()
        {
            if (loadedCustomUIs == null || loadedCustomUIs.Count == 0)
            {
                // Nothing to do.
                return;
            }
            Application excelApp = new Application(ExcelDnaUtil.Application);
            foreach (XmlNode xmlCustomUI in loadedCustomUIs)
            {
                try
                {
                    RemoveCommandBarControls(excelApp, xmlCustomUI.ChildNodes);
                }
                catch (Exception e)
                {
                    // Suppress exceptions
                    Debug.Print("ExcelCommandBars: Error removing controls: {0}", e);
                }

            }
            loadedCustomUIs.Clear();
        }

        private static void AddCommandBarControls(Application excelApp, XmlNodeList xmlNodes, GetImageDelegate getImage)
        {
            foreach (XmlNode childNode in xmlNodes)
            {
                if (childNode.Name == "commandBar")
                {
                    string barName;
                    CommandBar bar = GetCommandBarFromIdOrName(excelApp, childNode.Attributes, out barName);
                    if (bar != null)
                    {
                        AddControls(bar.Controls, childNode.ChildNodes, getImage);
                    }
                    else
                    {
                        MsoBarPosition barPosition = MsoBarPosition.msoBarLeft;
                        XmlAttribute posAttribute = childNode.Attributes["position"];
                        if (posAttribute == null)
                        {
                            // Compatible with original patch
                            posAttribute = childNode.Attributes["MsoBarPosition"];
                        }
                        if (posAttribute != null)
                        {
                            if (Enum.IsDefined(typeof(MsoBarPosition), posAttribute.Value))
                                barPosition = (MsoBarPosition)Enum.Parse(typeof(MsoBarPosition), posAttribute.Value, false);
                        }

                        bar = excelApp.CommandBars.Add(barName, barPosition);
                        AddControls(bar.Controls, childNode.ChildNodes, getImage);
                    }

                }
            }
        }

        private static void RemoveCommandBarControls(Application excelApp, XmlNodeList xmlNodes)
        {
            foreach (XmlNode childNode in xmlNodes)
            {
                if (childNode.Name == "commandBar")
                {
                    string barName;
                    CommandBar bar = GetCommandBarFromIdOrName(excelApp, childNode.Attributes, out barName);
                    if (bar != null)
                    {
                        RemoveControls(bar.Controls, childNode.ChildNodes);

                        if (bar.Controls.Count() == 0)
                        {
                            bar.Delete();
                        }
                    }
                }
            }
        }

        // This method contributed by Benoit Patra (see GitHub pull request: https://github.com/Excel-DNA/Excel-DNA/pull/1)
        // TODO: Still need to sort out the Id property
        //       This version is temporary, should behave the same as v.0.32
        private static CommandBar GetCommandBarFromIdOrName(Application excelApp, XmlAttributeCollection nodeAttributes, out string barName)
        {
            XmlAttribute name = nodeAttributes["name"];
            if (name == null) throw new ArgumentException("CommandBar attributes must contain name");
            barName = name.Value;

            CommandBar bar = null;
            for (int i = 1; i <= excelApp.CommandBars.Count; i++)
            {
                if (excelApp.CommandBars[i].Name == barName)
                {
                    bar = excelApp.CommandBars[i];
                    break;
                }
            }
            return bar;
        }

        //// We cannot rely only on name to recover the proper CommandBar so we have the possibility to use the ID (which is used in priority).
        //// Indeed there are two CommandBar for "Cell" see  http://msdn.microsoft.com/en-us/library/office/gg469862(v=office.14).aspx
        //// However, at the time of the writing there is a mistake: "Application.CommandBars(Application.CommandBars("Cell").Index + 3)" is false in practice
        //private static CommandBar GetCommandBarFromIdOrName(Application excelApp,XmlAttributeCollection nodeAttributes, out string barName)
        //{
        //    var id =  nodeAttributes["id"];
        //    var name = nodeAttributes["name"];
        //    if(name ==null) throw new ArgumentException("commandBar attributes must contain name");

        //    barName = name.Value;
        //    if (id != null)
        //    {
        //        string barId = id.Value;
        //        CommandBar bar = null;
        //        for (int i = 1; i <= excelApp.CommandBars.Count; i++)
        //        {
        //            if (excelApp.CommandBars[i].Id == barId)
        //            {
        //                bar = excelApp.CommandBars[i];
        //                break;
        //            }
        //        }
        //        return bar;
        //    }
        //    else
        //    {

        //        CommandBar bar = null;
        //        for (int i = 1; i <= excelApp.CommandBars.Count; i++)
        //        {
        //            if (excelApp.CommandBars[i].Name == barName)
        //            {
        //                bar = excelApp.CommandBars[i];
        //                break;
        //            }
        //        }
        //        return bar;
        //    }
        //}


        private static void AddControls(CommandBarControls parentControls, XmlNodeList xmlNodes, GetImageDelegate getImage)
        {
            foreach (XmlNode childNode in xmlNodes)
            {
                AddControl(parentControls, childNode, getImage);
            }
        }

        private static void RemoveControls(CommandBarControls parentControls, XmlNodeList xmlNodes)
        {
            foreach (XmlNode childNode in xmlNodes)
            {
                RemoveControl(parentControls, childNode);
            }
        }

        private static void AddControl(CommandBarControls parentControls, XmlNode xmlNode, GetImageDelegate getImage)
        {
            if (xmlNode.Name == "popup")
            {
                string controlName = xmlNode.Attributes["caption"].Value;
                object before = ReadControlBeforeAttribute(xmlNode);
                CommandBarPopup newPopup = parentControls.AddPopup(controlName, before);
                ApplyControlAttributes(newPopup, xmlNode, getImage);
                AddControls(newPopup.Controls, xmlNode.ChildNodes, getImage);
            }
            else if (xmlNode.Name == "button")
            {
                object before = ReadControlBeforeAttribute(xmlNode);
                CommandBarButton newButton = parentControls.AddButton(before);
                ApplyControlAttributes(newButton, xmlNode, getImage);
            }
        }

        private static void RemoveControl(CommandBarControls parentControls, XmlNode xmlNode)
        {
            if (xmlNode.Name == "popup")
            {
                string controlName = xmlNode.Attributes["caption"].Value;
                CommandBarPopup cb = (parentControls[controlName] as CommandBarPopup);
                if (cb != null)
                {
                    RemoveControls(cb.Controls, xmlNode.ChildNodes);
                    if (cb.Controls.Count() == 0)
                    {
                        cb.Delete(true);
                    }
                }
            }
            if (xmlNode.Name == "button")
            {
                string controlName = xmlNode.Attributes["caption"].Value;
                parentControls[controlName].Delete(true);
            }
        }

        private static object ReadControlBeforeAttribute(XmlNode xmlNode)
        {
            // Before is one of the parameters of CommandBarControls.Add. If not set, we'll pass Missing.Value.
            // We also allow this to be a string... e.g. "Help"
            object before = Missing.Value;
            XmlAttribute beforeAttribute = xmlNode.Attributes["before"];
            if (beforeAttribute != null && !string.IsNullOrEmpty(beforeAttribute.Value))
            {
                int beforeValue;
                if (int.TryParse(beforeAttribute.Value, out beforeValue))
                {
                    return beforeValue;
                }
                else
                {
                    // Assume it is a string referring to a control
                    return beforeAttribute.Value;
                }
            }
            return before;
        }

        private static void ApplyControlAttributes(CommandBarControl control, XmlNode xmlNode, GetImageDelegate getImage)
        {
            foreach (XmlAttribute att in xmlNode.Attributes)
            {
                ApplyControlAttribute(control, att.Name, att.Value, getImage);
            }
        }

        private static void ApplyControlAttribute(CommandBarControl control, string attribute, string value, GetImageDelegate getImage)
        {
            switch (attribute)
            {
                case "caption":
                    control.Caption = value;
                    break;
                case "height":
                    int height;
                    if (int.TryParse(value, out height))
                    {
                        control.Height = height;
                    }
                    else
                    {
                        Debug.Print("Could not parse 'height' attribute: {0}", value);
                    }
                    break;
                case "onAction":
                    control.OnAction = value;
                    break;
                case "enabled":
                    bool enabled;
                    if (bool.TryParse(value, out enabled))
                    {
                        control.Enabled = enabled;
                    }
                    else
                    {
                        Debug.Print("Could not parse 'enabled' attribute: {0}", value);
                    }
                    break;
                case "beginGroup":
                    bool beginGroup;
                    if (bool.TryParse(value, out beginGroup))
                    {
                        control.BeginGroup = beginGroup;
                    }
                    else
                    {
                        Debug.Print("Could not parse 'beginGroup' attribute: {0}", value);
                    }
                    break;
                case "helpFile":
                    control.HelpFile = value;
                    break;
                case "helpContextId":
                    int helpContextId;
                    if (int.TryParse(value, out helpContextId))
                    {
                        control.HelpContextId = helpContextId;
                    }
                    else
                    {
                        Debug.Print("Could not parse 'helpContextId' attribute: {0}", value);
                    }
                    break;
                case "tag":
                    control.Tag = value;
                    break;
                case "tooltipText":
                    control.TooltipText = value;
                    break;
                case "shortcutText":
                    if (control is CommandBarButton)
                    {
                        (control as CommandBarButton).ShortcutText = value;
                    }
                    else
                    {
                        Debug.Print("shortcutText only supported on Buttons");
                    }
                    break;
                case "faceId":
                    if (control is CommandBarButton)
                    {
                        int faceId;
                        if (int.TryParse(value, out faceId))
                        {
                            (control as CommandBarButton).FaceId = faceId;
                        }
                        else
                        {
                            Debug.Print("Could not parse 'faceId' attribute: {0}", value);
                        }
                    }
                    else
                    {
                        Debug.Print("faceId only supported on Buttons");
                    }
                    break;
                case "image":
                    if (control is CommandBarButton)
                    {
                        Bitmap image = getImage(value);
                        if (image != null)
                        {
                            (control as CommandBarButton).SetButtonImage(image);
                        }
                        else
                        {
                            Debug.Print("Could not find or load image {0}", value);
                        }
                    }
                    else
                    {
                        Debug.Print("image only supported on Buttons");
                    }
                    break;
                case "style":
                case "MsoButtonStyle":  // Compatible with original style code.
                    if (control is CommandBarButton)
                    {
                        if (Enum.IsDefined(typeof(MsoButtonStyle), value))
                            (control as CommandBarButton).Style = (MsoButtonStyle)Enum.Parse(typeof(MsoButtonStyle), value, false);
                        else
                            (control as CommandBarButton).Style = MsoButtonStyle.msoButtonAutomatic;
                    }
                    else
                    {
                        Debug.Print("style only supported on Buttons");
                    }
                    break;
                default:
                    Debug.Print("Unknown attribute '{0}' - ignoring.", attribute);
                    break;
            }
        }

        // Some minimal wrappers for the office types.
        private class Application
        {
            object _object;
            Type _type;

            public Application(object application)
            {
                _object = application;
                _type = _object.GetType();
            }

            public CommandBars CommandBars
            {
                get
                {
                    object commandBars = _type.InvokeMember("CommandBars", BindingFlags.GetProperty, null, _object, null);
                    return new CommandBars(commandBars);
                }
            }
        }
    }

    // Pattern for event handlers: myControl.GetType().GetEvent("Click").AddEventHandler(myControl, myControlHandler);
    // - Nope - need to do explicit ConnectionPoints.... http://www.codeproject.com/KB/cs/zetalatebindingcomevents.aspx

    // CommandBar events info: http://msdn.microsoft.com/en-us/library/aa189726(office.10).aspx
    /*
        The CommandBars collection and the CommandBarButton and CommandBarComboBox objects expose the following event procedures that you can use to run code in response to an event:

        The CommandBars collection supports the OnUpdate event, which is triggered in response to changes made to a Microsoft® Office document that might affect the state of any visible command bar or command bar control. For example, the OnUpdate event occurs when a user changes the selection in an Office document. You can use this event to change the availability or state of command bars or command bar controls in response to actions taken by the user.
        Note   The OnUpdate event can be triggered repeatedly in many different contexts. Any code you add to this event that does a lot of processing or performs a number of actions might affect the performance of your application.
        The CommandBarButton control exposes a Click event that is triggered when a user clicks a command bar button. You can use this event to run code when the user clicks a command bar button.
        The CommandBarComboBox control exposes a Change event that is triggered when a user makes a selection from a combo box control. You can use this method to take an action depending on what selection the user makes from a combo box control on a command bar.
     */

    public class CommandBar
    {
        object ComObject;
        Type ComObjectType;

        internal CommandBar(object commandBar)
        {
            ComObject = commandBar;
            ComObjectType = ComObject.GetType();
        }

        public object GetComObject()
        {
            return ComObject;
        }

        public CommandBarControls Controls
        {
            get
            {
                object controls = ComObjectType.InvokeMember("Controls", BindingFlags.GetProperty, null, ComObject, null);
                return new CommandBarControls(controls);
            }
        }

        public string Name
        {
            get
            {
                object controls = ComObjectType.InvokeMember("Name", BindingFlags.GetProperty, null, ComObject, null);
                return controls.ToString();
            }
        }

        public bool Visible
        {
            get
            {
                return (bool)ComObjectType.InvokeMember("Visible", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public CommandBarControl FindControl(object type, object id, object tag, object visible, object recursive)
        {
            object result = ComObjectType.InvokeMember("FindControl", BindingFlags.InvokeMethod, null, ComObject, new object[] { type, id, tag, visible, recursive });
            if (result == null) return null;
            return new CommandBarControl(result);
        }

        public void Delete()
        {
            ComObjectType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComObject, null);
        }

    }

    public class CommandBars
    {
        object _object;
        Type _type;

        internal CommandBars(object commandBars)
        {
            _object = commandBars;
            _type = _object.GetType();
        }

        public object GetComObject()
        {
            return _object;
        }

        public CommandBar Add(string name, MsoBarPosition barPosition)
        {
            object commandBar = _type.InvokeMember("Add", BindingFlags.InvokeMethod, null, _object, new object[] { name, barPosition, Type.Missing, true });
            CommandBar cb = new CommandBar(commandBar);
            cb.Visible = true;
            return new CommandBar(commandBar);
        }

        public CommandBar this[string name]
        {
            get
            {
                object commandBar = _type.InvokeMember("", BindingFlags.GetProperty, null, _object, new object[] { name });
                return new CommandBar(commandBar);
            }
        }

        public CommandBar this[int i]
        {
            get
            {
                object commandBar = _type.InvokeMember("", BindingFlags.GetProperty, null, _object, new object[] { i });
                return new CommandBar(commandBar);
            }
        }

        public int Count
        {
            get
            {
                object i = _type.InvokeMember("Count", BindingFlags.GetProperty, null, _object, null);
                return Convert.ToInt32(i);
            }
        }

        //public event EventHandler OnUpdate
        //{
        //    add
        //    {
        //    }
        //    remove
        //    {
        //    }
        //}
    }

    public class CommandBarControl
    {
        private static Guid guidCommandBarButton = new Guid("000C030E-0000-0000-C000-000000000046");
        private static Guid guidCommandBarPopup = new Guid("000C030A-0000-0000-C000-000000000046");
        private static Guid guidCommandBarComboBox = new Guid("000C030C-0000-0000-C000-000000000046");

        internal protected object ComObject;
        internal protected Type ComObjectType;

        internal CommandBarControl(object commandBarControl)
        {
            ComObject = commandBarControl;
            ComObjectType = ComObject.GetType();
        }

        internal static CommandBarControl CreateCommandBarControl(MsoControlType controlType, object commandBarControl)
        {
            if (controlType == MsoControlType.msoControlButton)
            {
                return new CommandBarButton(commandBarControl);
            }
            if (controlType == MsoControlType.msoControlPopup)
            {
                return new CommandBarPopup(commandBarControl);
            }
            if (controlType == MsoControlType.msoControlComboBox)
            {
                return new CommandBarComboBox(commandBarControl);
            }
            return new CommandBarControl(commandBarControl);
        }

        // In this case we check the interfaces for the right type
        internal static CommandBarControl CreateCommandBarControl(object commandBarControl)
        {
            IntPtr pUnk = Marshal.GetIUnknownForObject(commandBarControl);

            IntPtr pButton;
            Marshal.QueryInterface(pUnk, ref guidCommandBarButton, out pButton);
            if (pButton != IntPtr.Zero)
            {
                return new CommandBarButton(commandBarControl);
            }

            IntPtr pPopup;
            Marshal.QueryInterface(pUnk, ref guidCommandBarPopup, out pPopup);
            if (pPopup != IntPtr.Zero)
            {
                return new CommandBarPopup(commandBarControl);
            }

            IntPtr pComboBox;
            Marshal.QueryInterface(pUnk, ref guidCommandBarPopup, out pComboBox);
            if (pComboBox != IntPtr.Zero)
            {
                return new CommandBarComboBox(commandBarControl);
            }

            return new CommandBarControl(commandBarControl);
        }


        public object GetComObject()
        {
            return ComObject;
        }

        public string Caption
        {
            get
            {
                return (string)ComObjectType.InvokeMember("Caption", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("Caption", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public string Tag
        {
            get
            {
                return (string)ComObjectType.InvokeMember("Tag", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("Tag", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public string TooltipText
        {
            get
            {
                return (string)ComObjectType.InvokeMember("TooltipText", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("TooltipText", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public string OnAction
        {
            get
            {
                return (string)ComObjectType.InvokeMember("OnAction", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("OnAction", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public bool BeginGroup
        {
            get
            {
                return (bool)ComObjectType.InvokeMember("BeginGroup", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("BeginGroup", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public bool Enabled
        {
            get
            {
                return (bool)ComObjectType.InvokeMember("Enabled", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("Enabled", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public int Height
        {
            get
            {
                return (int)ComObjectType.InvokeMember("Height", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("Height", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public string HelpFile
        {
            get
            {
                return (string)ComObjectType.InvokeMember("HelpFile", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("HelpFile", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public int HelpContextId
        {
            get
            {
                return (int)ComObjectType.InvokeMember("HelpContextId", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("HelpContextId", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public bool Visible
        {
            get
            {
                return (bool)ComObjectType.InvokeMember("Visible", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public int Index
        {
            get
            {
                return (int)ComObjectType.InvokeMember("Index", BindingFlags.GetProperty, null, ComObject, null);
            }
        }

        public void Delete(object Temporary)
        {
            ComObjectType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComObject, new object[] { Temporary });
        }
    }

    public class CommandBarControls
    {
        object ComObject;
        Type ComObjectTtpe;

        internal CommandBarControls(object commandBarControls)
        {
            ComObject = commandBarControls;
            ComObjectTtpe = ComObject.GetType();
        }

        public object GetComObject()
        {
            return ComObject;
        }

        public CommandBarControl this[string name]
        {
            get
            {
                object commandBarControl = ComObjectTtpe.InvokeMember("", BindingFlags.GetProperty, null, ComObject, new object[] { name });
                return CommandBarControl.CreateCommandBarControl(commandBarControl);
            }
        }

        public CommandBarControl this[int id]
        {
            get
            {
                object commandBarControl = ComObjectTtpe.InvokeMember(
                    "", BindingFlags.GetProperty, null, ComObject, new object[] { id });
                return CommandBarControl.CreateCommandBarControl(commandBarControl);
            }
        }

        public int Count()
        {
            object i = ComObjectTtpe.InvokeMember("Count", BindingFlags.GetProperty, null, ComObject, null);
            return Convert.ToInt32(i);
        }

        public CommandBarControl Add(MsoControlType controlType, object Id, object Parameter, object Before, object Temporary)
        {
            return FindOrAdd(controlType, null, Id, Parameter, Before, Temporary);
        }

        internal CommandBarControl FindOrAdd(MsoControlType controlType, string name, object Id, object Parameter, object Before, object Temporary)
        {
            if (name != null)
            {
                // Try to find an existing control with this name
                string findName = name.Replace("&", "");
                for (int i = 1; i <= Count(); i++)
                {
                    CommandBarControl control = this[i];
                    string caption = control.Caption;
                    if (!String.IsNullOrEmpty(caption))
                        if (caption.Replace("&", "") == findName)
                            return control;
                }
            }

            object /*CommandBarControl*/ newControl = ComObjectTtpe.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComObject,
                new object[] { controlType, Id, Parameter, Before, Temporary });

            return CommandBarControl.CreateCommandBarControl(controlType, newControl);
        }

        // Normalizes the before value.
        // If before is missing or an int, already OK.
        // If it is a string, find the control in this collection and return its index.
        private object FindControlIndexBefore(object before /* should be int or Missing or string referring to control in this collection */)
        {
            if (before is Missing || before is int)
            {
                return before;
            }

            object beforeIndex = Missing.Value;
            if (before is string)
            {
                for (int i = 1; i <= Count(); i++)
                {
                    CommandBarControl control = this[i];
                    string caption = control.Caption;
                    if (!String.IsNullOrEmpty(caption))
                    {
                        if (caption.Replace("&", "") == ((string)before).Replace("&", ""))
                        {
                            // This is the one!
                            beforeIndex = i;
                            break;
                        }
                    }
                }
            }
            return beforeIndex;
        }

        public CommandBarButton AddButton()
        {
            return AddButton(Type.Missing);
        }

        // before should be int or Missing or string referring to control in this collection
        public CommandBarButton AddButton(object before)
        {
            object beforeIndex = FindControlIndexBefore(before);
            return (CommandBarButton)Add(MsoControlType.msoControlButton, 1, Type.Missing, beforeIndex, true);
        }

        public CommandBarPopup AddPopup(string name)
        {
            return AddPopup(name, Type.Missing);
        }

        // before should be int or Missing or string referring to control in this collection
        public CommandBarPopup AddPopup(string name, object before)
        {
            object beforeIndex = FindControlIndexBefore(before);
            return (CommandBarPopup)FindOrAdd(MsoControlType.msoControlPopup, name, 1, Type.Missing, beforeIndex, true);
        }


        public CommandBarComboBox AddComboBox()
        {
            return AddComboBox(Type.Missing);
        }

        // before should be int or Missing or string referring to control in this collection
        public CommandBarComboBox AddComboBox(object before)
        {
            object beforeIndex = FindControlIndexBefore(before);
            return (CommandBarComboBox)Add(MsoControlType.msoControlComboBox, 1, Type.Missing, beforeIndex, true);
        }

        private void Remove(MsoControlType controlType, object id)
        {
            for (int i = 1; i <= Count(); i++)
            {
                if (!String.IsNullOrEmpty(this[i].Caption))
                    if (this[i].Caption.Replace("&", "") == id.ToString().Replace("&", ""))
                        this[i].Delete(true);
            }
        }

        public void RemoveButton()
        {
            Remove(MsoControlType.msoControlButton, 1);
        }

        public void RemovePopup(string name)
        {
            Remove(MsoControlType.msoControlPopup, name);
        }
    }

    public class CommandBarButton : CommandBarControl
    {
        internal CommandBarButton(object commandBarCom)
            : base(commandBarCom)
        {
        }

        public void SetButtonImage(Bitmap buttonImage)
        {
            // TODO: Consider using Picture property for Excel 2002+ (and Mask?)
            //       http://support.microsoft.com/kb/286460?wa=wsignin1.0
            //       Remember that .NET Bitmap already implements IPictureDisp.

            // IDataObject oldContent = Clipboard.GetDataObject();
            Clipboard.SetImage(buttonImage);
            Type t = ComObject.GetType();
            t.InvokeMember("Style", BindingFlags.SetProperty, null, ComObject, new object[] { MsoButtonStyle.msoButtonIconAndCaption });
            t.InvokeMember("PasteFace", BindingFlags.InvokeMethod, null, ComObject, null);
            Clipboard.Clear();
            // Clipboard.SetDataObject(oldContent);
        }

        public int FaceId
        {
            get
            {
                return (int)ComObjectType.InvokeMember("FaceId", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("FaceId", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public MsoButtonStyle Style
        {
            get
            {
                return (MsoButtonStyle)ComObjectType.InvokeMember("Style", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("Style", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        public string ShortcutText
        {
            get
            {
                return (string)ComObjectType.InvokeMember("ShortcutText", BindingFlags.GetProperty, null, ComObject, null);
            }
            set
            {
                ComObjectType.InvokeMember("ShortcutText", BindingFlags.SetProperty, null, ComObject, new object[] { value });
            }
        }

        // TODO: Decide whether to implement late-bound event handlers.
        //       Under .NET 4 with the Embed Interop Types option, it might not make sense to expand the late-bound wrappers any further.

        //public class CommandBarButtonClickEventArgs : EventArgs
        //{
        //    public bool CancelDefault;
        //}

        //public event EventHandler<CommandBarButtonClickEventArgs> Click
        //{
        //    add
        //    {
        //    }
        //    remove
        //    {
        //    }
        //}
    }

    public class CommandBarPopup : CommandBarControl
    {
        public CommandBarPopup(object commandBarCom)
            : base(commandBarCom)
        {
        }

        public CommandBarControls Controls
        {
            get
            {
                object controls = ComObjectType.InvokeMember("Controls", BindingFlags.GetProperty, null, ComObject, null);
                return new CommandBarControls(controls);
            }
        }
    }

    public class CommandBarComboBox : CommandBarControl
    {
        public CommandBarComboBox(object commandBarCom)
            : base(commandBarCom)
        {
        }
    }


    //#region Assembly Office.dll, v1.1.4322
    // C:\WINDOWS\assembly\GAC\Office\12.0.0.0__71e9bce111e9429c\Office.dll
    //#endregion

    //namespace Microsoft.Office.Core
    public enum MsoBarPosition
    {
        msoBarLeft = 0,
        msoBarTop = 1,
        msoBarRight = 2,
        msoBarBottom = 3,
        msoBarFloating = 4,
        msoBarPopup = 5,
        msoBarMenuBar = 6
    }

    public enum MsoButtonStyle
    {
        msoButtonAutomatic = 0,
        msoButtonIcon = 1,
        msoButtonCaption = 2,
        msoButtonIconAndCaption = 3,
        msoButtonIconAndWrapCaption = 7,
        msoButtonIconAndCaptionBelow = 11,
        msoButtonWrapCaption = 14,
        msoButtonIconAndWrapCaptionBelow = 15,
    }

    public enum MsoControlType
    {
        msoControlCustom = 0,
        msoControlButton = 1,
        msoControlEdit = 2,
        msoControlDropdown = 3,
        msoControlComboBox = 4,
        msoControlButtonDropdown = 5,
        msoControlSplitDropdown = 6,
        msoControlOCXDropdown = 7,
        msoControlGenericDropdown = 8,
        msoControlGraphicDropdown = 9,
        msoControlPopup = 10,
        msoControlGraphicPopup = 11,
        msoControlButtonPopup = 12,
        msoControlSplitButtonPopup = 13,
        msoControlSplitButtonMRUPopup = 14,
        msoControlLabel = 15,
        msoControlExpandingGrid = 16,
        msoControlSplitExpandingGrid = 17,
        msoControlGrid = 18,
        msoControlGauge = 19,
        msoControlGraphicCombo = 20,
        msoControlPane = 21,
        msoControlActiveX = 22,
        msoControlSpinner = 23,
        msoControlLabelEx = 24,
        msoControlWorkPane = 25,
        msoControlAutoCompleteCombo = 26,
    }
}
