using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using System.Diagnostics;

using ExcelDna.Serialization;
using System.Xml.Serialization;
using System.IO;

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
            using (var sr = new StringReader(dnaLibraryWrapper))
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
                    string barName = childNode.Attributes["name"].Value;
                    CommandBar bar = null;
                    for (int i = 1; i <= excelApp.CommandBars.Count; i++)
                    {
                        if (excelApp.CommandBars[i].Name == barName)
                        {
                            bar = excelApp.CommandBars[i];
                            break;
                        }
                    }
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
                        if ( posAttribute != null)
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
                    string barName = childNode.Attributes["name"].Value;
                    CommandBar bar = null;
                    for (int i = 1; i <= excelApp.CommandBars.Count; i++)
                    {
                        if (excelApp.CommandBars[i].Name == barName)
                        {
                            bar = excelApp.CommandBars[i];
                            break;
                        }
                    }
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
            CommandBarControl newControl;
            if (xmlNode.Name == "popup")
            {
                string controlName = xmlNode.Attributes["caption"].Value;
                newControl = parentControls.AddPopup(controlName);
                ApplyControlAttributes(newControl, xmlNode, getImage);
                AddControls(newControl.Controls, xmlNode.ChildNodes, getImage);
            }
            else if (xmlNode.Name == "button")
            {
                newControl = parentControls.AddButton();
                ApplyControlAttributes(newControl, xmlNode, getImage);
            }
        }

        private static void RemoveControl(CommandBarControls parentControls, XmlNode xmlNode)
        {
            string controlName = xmlNode.Attributes["caption"].Value;
            if (xmlNode.Name == "popup")
            {
                CommandBarControl cb = parentControls[controlName];
                RemoveControls(cb.Controls, xmlNode.ChildNodes);

                if (cb.Controls.Count() == 0)
                {
                    cb.Delete(true);
                }
            }
            if (xmlNode.Name == "button")
            {
                parentControls[controlName].Delete(true);
            }
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
                    Bitmap image = getImage(value);
                    if (image != null)
                    {
                        control.SetButtonImage(image);
                    }
                    else
                    {
                        Debug.Print("Could not find or load image {0}", value);
                    }
                    break;
                case "style":
                case "MsoButtonStyle":  // Compatible with original style code.
                    if (Enum.IsDefined(typeof(MsoButtonStyle), value))
                        control.Style = (MsoButtonStyle)Enum.Parse(typeof(MsoButtonStyle), value, false);
                    else
                        control.Style = MsoButtonStyle.msoButtonAutomatic;
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

            //public CommandBarControl AddCommandBarMenu(string commandBarName, int? before)
            //{
            //    CommandBarControl cd;
            //    if (before == null)
            //    {
            //        cd = CommandBars[commandBarName].Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            //    }
            //    else
            //    {
            //        int b = (int)before;
            //        cd = CommandBars[commandBarName].Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, b, true);
            //    }
            //    return cd;
            //}
        }
    }
    public class CommandBar
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

        public string Name
        {
            get
            {
                object controls = _type.InvokeMember("Name", BindingFlags.GetProperty, null, _object, null);
                return controls.ToString();
            }
        }

        public bool Visible
        {
            set 
            {
                _type.InvokeMember("Visible", BindingFlags.SetProperty, null, _object, new object[] { value });
            }
        }


        public void Delete()
        {
            _type.InvokeMember("Delete", BindingFlags.InvokeMethod, null, _object, null);
        }

    }

    public class CommandBars
    {
        object _object;
        Type _type;

        public CommandBars(object commandBars)
        {
            _object = commandBars;
            _type = _object.GetType();
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

    }

    public class CommandBarControl
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

        public MsoButtonStyle Style
        {
            get
            {
                return (MsoButtonStyle)_type.InvokeMember("Style", BindingFlags.GetProperty, null, _object, null);
            }
            set
            {
                _type.InvokeMember("Style", BindingFlags.SetProperty, null, _object, new object[] { value });
            }
        }
            
        public void Delete(object Temporary)
        {
            _type.InvokeMember("Delete", BindingFlags.InvokeMethod, null, _object, new object[] { Temporary });
        }
    }

    public class CommandBarControls
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
                object commandBarControl = _type.InvokeMember("", BindingFlags.GetProperty, null, _object, new object[] { name });
                return new CommandBarControl(commandBarControl);
            }
        }

        public CommandBarControl this[int id]
        {
            get
            {
                object commandBarControl = _type.InvokeMember(
                    "", BindingFlags.GetProperty, null, _object, new object[] { id });
                return new CommandBarControl(commandBarControl);
            }
        }

        public int Count()
        {
            object i = _type.InvokeMember("Count", BindingFlags.GetProperty, null, _object, null);
            return Convert.ToInt32(i);
        }

        private CommandBarControl Add(MsoControlType controlType, string name, object Id, object Parameter, object Before, object Temporary)
        {
            for (int i = 1; i <= Count(); i++)
            {
                if (!String.IsNullOrEmpty(this[i].Caption))
                    if (this[i].Caption.Replace("&", "") == name.Replace("&", ""))
                        return this[i];
            }

            object /*CommandBarControl*/ newControl = _type.InvokeMember("Add", BindingFlags.InvokeMethod, null, _object,
                new object[] { controlType, Id, Parameter, Before, Temporary });
            return new CommandBarControl(newControl);
        }

        public CommandBarControl AddButton()
        {
            return Add(MsoControlType.msoControlButton, "", 1, Type.Missing, Type.Missing, true);
        }

        public CommandBarControl AddPopup(string name)
        {
            return Add(MsoControlType.msoControlPopup, name, 1, Type.Missing, Type.Missing, true);
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

}
