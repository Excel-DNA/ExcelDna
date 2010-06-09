using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using Microsoft.Office.Core;
using System.Diagnostics; // Not from PIA, but from ComInterop.cs

namespace ExcelDna.Integration.CustomUI
{
    internal static class ExcelCommandBars
    {
        static List<XmlNode> loadedCustomUIs = new List<XmlNode>();
        public static void LoadCommandBars(XmlNode xmlCustomUI, DnaLibrary dnaLibrary)
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
                AddCommandBarControls(excelApp, xmlCustomUI.ChildNodes, dnaLibrary);
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

        private static void AddCommandBarControls(Application excelApp, XmlNodeList xmlNodes, DnaLibrary dnaLibrary)
        {
            foreach (XmlNode childNode in xmlNodes)
            {
                if (childNode.Name == "commandBar")
                {
                    string barName = childNode.Attributes["name"].Value;
                    CommandBar bar = excelApp.CommandBars[barName];
                    if (bar != null)
                    {
                        AddControls(bar.Controls, childNode.ChildNodes, dnaLibrary);
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
                    string barName = childNode.Attributes["name"].Value;
                    CommandBar bar = excelApp.CommandBars[barName];
                    if (bar != null)
                    {
                        RemoveControls(bar.Controls, childNode.ChildNodes);
                    }
                }
            }
        }

        private static void AddControls(CommandBarControls parentControls, XmlNodeList xmlNodes, DnaLibrary dnaLibrary)
        {
            foreach (XmlNode childNode in xmlNodes)
            {
                AddControl(parentControls, childNode, dnaLibrary);
            }
        }

        private static void RemoveControls(CommandBarControls parentControls, XmlNodeList xmlNodes)
        {
            foreach (XmlNode childNode in xmlNodes)
            {
                RemoveControl(parentControls, childNode);
            }
        }

        private static void AddControl(CommandBarControls parentControls, XmlNode xmlNode, DnaLibrary dnaLibrary)
        {
            CommandBarControl newControl;
            if (xmlNode.Name == "popup")
            {
                newControl = parentControls.AddPopup();
                ApplyControlAttributes(newControl, xmlNode, dnaLibrary);
                AddControls(newControl.Controls, xmlNode.ChildNodes, dnaLibrary);
            }
            else if (xmlNode.Name == "button")
            {
                newControl = parentControls.AddButton();
                ApplyControlAttributes(newControl, xmlNode, dnaLibrary);
            }
        }

        private static void RemoveControl(CommandBarControls parentControls, XmlNode xmlNode)
        {
            // top level controls only - no recursing down.
            if (xmlNode.Name == "popup" || xmlNode.Name == "button")
            {
                string controlName = xmlNode.Attributes["caption"].Value;
                parentControls[controlName].Delete(true);
            }
        }

        private static void ApplyControlAttributes(CommandBarControl control, XmlNode xmlNode, DnaLibrary dnaLibrary)
        {
            foreach (XmlAttribute att in xmlNode.Attributes)
            {
                ApplyControlAttribute(control, att.Name, att.Value, dnaLibrary);
            }
        }

        private static void ApplyControlAttribute(CommandBarControl control, string attribute, string value, DnaLibrary dnaLibrary)
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
                    control.ShortcutText = value;
                    break;
                case "faceId":
                    int faceId;
                    if (int.TryParse(value, out faceId))
                    {
                        control.FaceId = faceId;
                    }
                    else
                    {
                        Debug.Print("Could not parse 'faceId' attribute: {0}", value);
                    }
                    break;
                case "image":
                    Bitmap image = dnaLibrary.GetImage(value);
                    if (image != null)
                    {
                        control.SetButtonImage(image);
                    }
                    else
                    {
                        Debug.Print("Could not find or load image {0}", value);
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

            public CommandBarControl AddCommandBarMenu(string commandBarName, int? before)
            {
                CommandBarControl cd;
                if (before == null)
                {
                    cd = CommandBars[commandBarName].Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
                }
                else
                {
                    int b = (int)before;
                    cd = CommandBars[commandBarName].Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, b, true);
                }
                return cd;
            }
        }

        private class CommandBar
        {
            object _object;
            Type _type;
            public CommandBar(object commandBar)
            {
                _object = commandBar;
                _type = _object.GetType();
            }

            public CommandBarControls Controls
            {
                get
                {
                    object controls = _type.InvokeMember("Controls", BindingFlags.GetProperty, null, _object, null);
                    return new CommandBarControls(controls);
                }
            }
        }

        private class CommandBars
        {
            object _object;
            Type _type;
            public CommandBars(object commandBars)
            {
                _object = commandBars;
                _type = _object.GetType();
            }

            public CommandBar this[string name]
            {
                get
                {
                    object commandBar = _type.InvokeMember("", BindingFlags.GetProperty, null, _object, new object[] { name });
                    return new CommandBar(commandBar);
                }
            }
        }

        private class CommandBarControl
        {
            object _object;
            Type _type;
            public CommandBarControl(object commandBarControl)
            {
                _object = commandBarControl;
                _type = _object.GetType();
            }

            public CommandBarControls Controls
            {
                get
                {
                    object controls = _type.InvokeMember("Controls", BindingFlags.GetProperty, null, _object, null);
                    return new CommandBarControls(controls);
                }
            }

            public void SetButtonImage(Bitmap buttonImage)
            {
                Clipboard.SetImage(buttonImage);
                Type t = _object.GetType();
                t.InvokeMember("Style", BindingFlags.SetProperty, null, _object, new object[] { MsoButtonStyle.msoButtonIconAndCaption });
                t.InvokeMember("PasteFace", BindingFlags.InvokeMethod, null, _object, null);
            }

            public int FaceId
            {
                get
                {
                    return (int)_type.InvokeMember("FaceId", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("FaceId", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public string Caption
            {
                get
                {
                    return (string)_type.InvokeMember("Caption", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("Caption", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public string Tag
            {
                get
                {
                    return (string)_type.InvokeMember("Tag", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("Tag", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public string ShortcutText
            {
                get
                {
                    return (string)_type.InvokeMember("ShortcutText", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("ShortcutText", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public string TooltipText
            {
                get
                {
                    return (string)_type.InvokeMember("TooltipText", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("TooltipText", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public string OnAction
            {
                get
                {
                    return (string)_type.InvokeMember("OnAction", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("OnAction", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public bool BeginGroup
            {
                get
                {
                    return (bool)_type.InvokeMember("BeginGroup", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("BeginGroup", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public bool Enabled
            {
                get
                {
                    return (bool)_type.InvokeMember("Enabled", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("Enabled", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public int Height
            {
                get
                {
                    return (int)_type.InvokeMember("Height", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("Height", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public string HelpFile
            {
                get
                {
                    return (string)_type.InvokeMember("HelpFile", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("HelpFile", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }

            public int HelpContextId
            {
                get
                {
                    return (int)_type.InvokeMember("HelpContextId", BindingFlags.GetProperty, null, _object, null);
                }
                set
                {
                    _type.InvokeMember("HelpContextId", BindingFlags.SetProperty, null, _object, new object[] { value });
                }
            }
            
            public void Delete(object Temporary)
            {
                _type.InvokeMember("Delete", BindingFlags.InvokeMethod, null, _object, new object[] { Temporary });
            }
        }

        private class CommandBarControls
        {
            object _object;
            Type _type;
            public CommandBarControls(object commandBarControls)
            {
                _object = commandBarControls;
                _type = _object.GetType();
            }

            public CommandBarControl this[string name]
            {
                get
                {
                    object commandBarControl = _type.InvokeMember(
                        "", BindingFlags.GetProperty, null, _object, new object[] { name });
                    return new CommandBarControl(commandBarControl);
                }
            }

            public CommandBarControl Add(MsoControlType controlType, object Id, object Parameter, object Before, object Temporary)
            {
                object /*CommandBarControl*/ newControl = _type.InvokeMember("Add", BindingFlags.InvokeMethod, null, _object,
                    new object[] { controlType, Id, Parameter, Before, Temporary });
                return new CommandBarControl(newControl);
            }

            public CommandBarControl AddButton()
            {
                return Add(MsoControlType.msoControlButton, 1, Type.Missing, Type.Missing, true);
            }

            public CommandBarControl AddPopup()
            {
                return Add(MsoControlType.msoControlPopup, 1, Type.Missing, Type.Missing, true);
            }
        }

    }
}
