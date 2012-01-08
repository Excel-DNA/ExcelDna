//#if _DYNAMIC_XMLSERIALIZER_COMPILATION
//[assembly:System.Security.AllowPartiallyTrustedCallers()]
//[assembly:System.Security.SecurityTransparent()]
//[assembly:System.Security.SecurityRules(System.Security.SecurityRuleSet.Level1)]
//#endif
//[assembly:System.Reflection.AssemblyVersionAttribute("0.30.4390.42893")]
//[assembly:System.Xml.Serialization.XmlSerializerVersionAttribute(ParentAssemblyId=@"6981d682-7dc2-4901-a5b8-9ede756a7433,", Version=@"4.0.0.0")]
//namespace Microsoft.Xml.Serialization.GeneratedAssembly {
namespace ExcelDna.Serialization
{

    public class XmlSerializationWriterDnaLibrary : System.Xml.Serialization.XmlSerializationWriter {

        public void Write8_DnaLibrary(object o) {
            WriteStartDocument();
            if (o == null) {
                WriteEmptyTag(@"DnaLibrary", @"");
                return;
            }
            TopLevelElement();
            Write7_DnaLibrary(@"DnaLibrary", @"", ((global::ExcelDna.Integration.DnaLibrary)o), false, false);
        }

        void Write7_DnaLibrary(string n, string ns, global::ExcelDna.Integration.DnaLibrary o, bool isNullable, bool needType) {
            if ((object)o == null) {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType) {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.DnaLibrary)) {
                }
                else {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"RuntimeVersion", @"", ((global::System.String)o.@RuntimeVersion));
            WriteAttribute(@"ShadowCopyFiles", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ShadowCopyFiles)));
            WriteAttribute(@"CreateSandboxedAppDomain", @"", ((global::System.String)o.@CreateSandboxedAppDomain));
            WriteAttribute(@"Language", @"", ((global::System.String)o.@Language));
            WriteAttribute(@"CompilerVersion", @"", ((global::System.String)o.@CompilerVersion));
            WriteAttribute(@"DefaultReferences", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultReferences)));
            WriteAttribute(@"DefaultImports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultImports)));
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.ExternalLibrary> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.ExternalLibrary>)o.@ExternalLibraries;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write2_ExternalLibrary(@"ExternalLibrary", @"", ((global::ExcelDna.Integration.ExternalLibrary)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Project> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Project>)o.@Projects;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write5_Project(@"Project", @"", ((global::ExcelDna.Integration.Project)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write3_Reference(@"Reference", @"", ((global::ExcelDna.Integration.Reference)a[ia]), false, false);
                    }
                }
            }
            if ((object)(o.@Code) != null){
                WriteValue(((global::System.String)o.@Code));
            }
            {
                global::System.Collections.Generic.List<global::System.Xml.XmlNode> a = (global::System.Collections.Generic.List<global::System.Xml.XmlNode>)o.@CustomUIs;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        if ((((global::System.Xml.XmlNode)a[ia])) is System.Xml.XmlNode || ((global::System.Xml.XmlNode)a[ia]) == null) {
                            WriteElementLiteral((System.Xml.XmlNode)((global::System.Xml.XmlNode)a[ia]), @"CustomUI", @"", false, false);
                        }
                        else {
                            throw CreateInvalidAnyTypeException(((global::System.Xml.XmlNode)a[ia]));
                        }
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Image> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Image>)o.@Images;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write6_Image(@"Image", @"", ((global::ExcelDna.Integration.Image)a[ia]), false, false);
                    }
                }
            }
            WriteEndElement(o);
        }

        void Write6_Image(string n, string ns, global::ExcelDna.Integration.Image o, bool isNullable, bool needType) {
            if ((object)o == null) {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType) {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.Image)) {
                }
                else {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"Path", @"", ((global::System.String)o.@Path));
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
            WriteEndElement(o);
        }

        void Write3_Reference(string n, string ns, global::ExcelDna.Integration.Reference o, bool isNullable, bool needType) {
            if ((object)o == null) {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType) {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.Reference)) {
                }
                else {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"AssemblyPath", @"", ((global::System.String)o.@AssemblyPath));
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
            WriteAttribute(@"Path", @"", ((global::System.String)o.@Path));
            WriteEndElement(o);
        }

        void Write5_Project(string n, string ns, global::ExcelDna.Integration.Project o, bool isNullable, bool needType) {
            if ((object)o == null) {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType) {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.Project)) {
                }
                else {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"Language", @"", ((global::System.String)o.@Language));
            WriteAttribute(@"CompilerVersion", @"", ((global::System.String)o.@CompilerVersion));
            WriteAttribute(@"DefaultReferences", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultReferences)));
            WriteAttribute(@"DefaultImports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultImports)));
            WriteAttribute(@"ExplicitExports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ExplicitExports)));
            WriteAttribute(@"ComServer", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ComServer)));
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write3_Reference(@"Reference", @"", ((global::ExcelDna.Integration.Reference)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>)o.@SourceItems;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write4_SourceItem(@"SourceItem", @"", ((global::ExcelDna.Integration.SourceItem)a[ia]), false, false);
                    }
                }
            }
            if ((object)(o.@Code) != null){
                WriteValue(((global::System.String)o.@Code));
            }
            WriteEndElement(o);
        }

        void Write4_SourceItem(string n, string ns, global::ExcelDna.Integration.SourceItem o, bool isNullable, bool needType) {
            if ((object)o == null) {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType) {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.SourceItem)) {
                }
                else {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"Path", @"", ((global::System.String)o.@Path));
            if ((object)(o.@Code) != null){
                WriteValue(((global::System.String)o.@Code));
            }
            WriteEndElement(o);
        }

        void Write2_ExternalLibrary(string n, string ns, global::ExcelDna.Integration.ExternalLibrary o, bool isNullable, bool needType) {
            if ((object)o == null) {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType) {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.ExternalLibrary)) {
                }
                else {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Path", @"", ((global::System.String)o.@Path));
            WriteAttribute(@"TypeLibPath", @"", ((global::System.String)o.@TypeLibPath));
            WriteAttribute(@"ComServer", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ComServer)));
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
            WriteAttribute(@"LoadFromBytes", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@LoadFromBytes)));
            WriteAttribute(@"ExplicitExports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ExplicitExports)));
            WriteEndElement(o);
        }

        protected override void InitCallbacks() {
        }
    }

    public class XmlSerializationReaderDnaLibrary : System.Xml.Serialization.XmlSerializationReader {

        public object Read8_DnaLibrary() {
            object o = null;
            Reader.MoveToContent();
            if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (((object) Reader.LocalName == (object)id1_DnaLibrary && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o = Read7_DnaLibrary(false, true);
                }
                else {
                    throw CreateUnknownNodeException();
                }
            }
            else {
                UnknownNode(null, @":DnaLibrary");
            }
            return (object)o;
        }

        global::ExcelDna.Integration.DnaLibrary Read7_DnaLibrary(bool isNullable, bool checkType) {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType) {
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.DnaLibrary o;
            o = new global::ExcelDna.Integration.DnaLibrary();
            if ((object)(o.@ExternalLibraries) == null) o.@ExternalLibraries = new global::System.Collections.Generic.List<global::ExcelDna.Integration.ExternalLibrary>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.ExternalLibrary> a_0 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.ExternalLibrary>)o.@ExternalLibraries;
            if ((object)(o.@Projects) == null) o.@Projects = new global::System.Collections.Generic.List<global::ExcelDna.Integration.Project>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Project> a_1 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Project>)o.@Projects;
            if ((object)(o.@References) == null) o.@References = new global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a_6 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
            if ((object)(o.@CustomUIs) == null) o.@CustomUIs = new global::System.Collections.Generic.List<global::System.Xml.XmlNode>();
            global::System.Collections.Generic.List<global::System.Xml.XmlNode> a_12 = (global::System.Collections.Generic.List<global::System.Xml.XmlNode>)o.@CustomUIs;
            if ((object)(o.@Images) == null) o.@Images = new global::System.Collections.Generic.List<global::ExcelDna.Integration.Image>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Image> a_13 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Image>)o.@Images;
            bool[] paramsRead = new bool[14];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[2] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[2] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id4_RuntimeVersion && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@RuntimeVersion = Reader.Value;
                    paramsRead[3] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id5_ShadowCopyFiles && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@ShadowCopyFiles = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id6_CreateSandboxedAppDomain && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@CreateSandboxedAppDomain = Reader.Value;
                    paramsRead[5] = true;
                }
                else if (!paramsRead[7] && ((object) Reader.LocalName == (object)id7_Language && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Language = Reader.Value;
                    paramsRead[7] = true;
                }
                else if (!paramsRead[8] && ((object) Reader.LocalName == (object)id8_CompilerVersion && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@CompilerVersion = Reader.Value;
                    paramsRead[8] = true;
                }
                else if (!paramsRead[9] && ((object) Reader.LocalName == (object)id9_DefaultReferences && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[9] = true;
                }
                else if (!paramsRead[10] && ((object) Reader.LocalName == (object)id10_DefaultImports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultImports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[10] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :RuntimeVersion, :ShadowCopyFiles, :CreateSandboxedAppDomain, :Language, :CompilerVersion, :DefaultReferences, :DefaultImports");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement) {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations0 = 0;
            int readerCount0 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None) {
                string tmp = null;
                if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                    if (((object) Reader.LocalName == (object)id11_ExternalLibrary && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_0) == null) Reader.Skip(); else a_0.Add(Read2_ExternalLibrary(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id12_Project && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_1) == null) Reader.Skip(); else a_1.Add(Read5_Project(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id13_Reference && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_6) == null) Reader.Skip(); else a_6.Add(Read3_Reference(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id14_CustomUI && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        a_12.Add((global::System.Xml.XmlNode)ReadXmlNode(true));
                    }
                    else if (((object) Reader.LocalName == (object)id15_Image && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_13) == null) Reader.Skip(); else a_13.Add(Read6_Image(false, true));
                    }
                    else {
                        UnknownNode((object)o, @":ExternalLibrary, :Project, :Reference, :CustomUI, :Image");
                    }
                }
                else if (Reader.NodeType == System.Xml.XmlNodeType.Text || 
                Reader.NodeType == System.Xml.XmlNodeType.CDATA || 
                Reader.NodeType == System.Xml.XmlNodeType.Whitespace || 
                Reader.NodeType == System.Xml.XmlNodeType.SignificantWhitespace) {
                    tmp = ReadString(tmp, false);
                    o.@Code = tmp;
                }
                else {
                    UnknownNode((object)o, @":ExternalLibrary, :Project, :Reference, :CustomUI, :Image");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations0, ref readerCount0);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.Image Read6_Image(bool isNullable, bool checkType) {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType) {
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.Image o;
            o = new global::ExcelDna.Integration.Image();
            bool[] paramsRead = new bool[3];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id16_Path && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id17_Pack && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[2] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :Path, :Pack");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement) {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations1 = 0;
            int readerCount1 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None) {
                if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                    UnknownNode((object)o, @"");
                }
                else {
                    UnknownNode((object)o, @"");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations1, ref readerCount1);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.Reference Read3_Reference(bool isNullable, bool checkType) {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType) {
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.Reference o;
            o = new global::ExcelDna.Integration.Reference();
            bool[] paramsRead = new bool[4];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id18_AssemblyPath && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@AssemblyPath = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id17_Pack && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[2] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id16_Path && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[3] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :AssemblyPath, :Pack, :Path");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement) {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations2 = 0;
            int readerCount2 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None) {
                if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                    UnknownNode((object)o, @"");
                }
                else {
                    UnknownNode((object)o, @"");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations2, ref readerCount2);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.Project Read5_Project(bool isNullable, bool checkType) {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType) {
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.Project o;
            try {
                o = (global::ExcelDna.Integration.Project)System.Activator.CreateInstance(typeof(global::ExcelDna.Integration.Project), System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.CreateInstance | System.Reflection.BindingFlags.NonPublic, null, new object[0], null);
            }
            catch (System.MissingMethodException) {
                throw CreateInaccessibleConstructorException(@"global::ExcelDna.Integration.Project");
            }
            catch (System.Security.SecurityException) {
                throw CreateCtorHasSecurityException(@"global::ExcelDna.Integration.Project");
            }
            if ((object)(o.@References) == null) o.@References = new global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a_3 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
            if ((object)(o.@SourceItems) == null) o.@SourceItems = new global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem> a_8 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>)o.@SourceItems;
            bool[] paramsRead = new bool[10];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id7_Language && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Language = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id8_CompilerVersion && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@CompilerVersion = Reader.Value;
                    paramsRead[2] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id9_DefaultReferences && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id10_DefaultImports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultImports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[5] = true;
                }
                else if (!paramsRead[6] && ((object) Reader.LocalName == (object)id19_ExplicitExports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@ExplicitExports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[6] = true;
                }
                else if (!paramsRead[7] && ((object) Reader.LocalName == (object)id20_ComServer && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@ComServer = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[7] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :Language, :CompilerVersion, :DefaultReferences, :DefaultImports, :ExplicitExports, :ComServer");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement) {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations3 = 0;
            int readerCount3 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None) {
                string tmp = null;
                if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                    if (((object) Reader.LocalName == (object)id13_Reference && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_3) == null) Reader.Skip(); else a_3.Add(Read3_Reference(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id21_SourceItem && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_8) == null) Reader.Skip(); else a_8.Add(Read4_SourceItem(false, true));
                    }
                    else {
                        UnknownNode((object)o, @":Reference, :SourceItem");
                    }
                }
                else if (Reader.NodeType == System.Xml.XmlNodeType.Text || 
                Reader.NodeType == System.Xml.XmlNodeType.CDATA || 
                Reader.NodeType == System.Xml.XmlNodeType.Whitespace || 
                Reader.NodeType == System.Xml.XmlNodeType.SignificantWhitespace) {
                    tmp = ReadString(tmp, false);
                    o.@Code = tmp;
                }
                else {
                    UnknownNode((object)o, @":Reference, :SourceItem");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations3, ref readerCount3);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.SourceItem Read4_SourceItem(bool isNullable, bool checkType) {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType) {
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.SourceItem o;
            o = new global::ExcelDna.Integration.SourceItem();
            bool[] paramsRead = new bool[4];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id17_Pack && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id16_Path && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[3] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Pack, :Name, :Path");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement) {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations4 = 0;
            int readerCount4 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None) {
                string tmp = null;
                if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                    UnknownNode((object)o, @"");
                }
                else if (Reader.NodeType == System.Xml.XmlNodeType.Text || 
                Reader.NodeType == System.Xml.XmlNodeType.CDATA || 
                Reader.NodeType == System.Xml.XmlNodeType.Whitespace || 
                Reader.NodeType == System.Xml.XmlNodeType.SignificantWhitespace) {
                    tmp = ReadString(tmp, false);
                    o.@Code = tmp;
                }
                else {
                    UnknownNode((object)o, @"");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations4, ref readerCount4);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.ExternalLibrary Read2_ExternalLibrary(bool isNullable, bool checkType) {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType) {
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.ExternalLibrary o;
            o = new global::ExcelDna.Integration.ExternalLibrary();
            bool[] paramsRead = new bool[6];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id16_Path && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id22_TypeLibPath && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@TypeLibPath = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id20_ComServer && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@ComServer = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[2] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id17_Pack && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[3] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id23_LoadFromBytes && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@LoadFromBytes = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id19_ExplicitExports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@ExplicitExports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[5] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Path, :TypeLibPath, :ComServer, :Pack, :LoadFromBytes, :ExplicitExports");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement) {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations5 = 0;
            int readerCount5 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None) {
                if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                    UnknownNode((object)o, @"");
                }
                else {
                    UnknownNode((object)o, @"");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations5, ref readerCount5);
            }
            ReadEndElement();
            return o;
        }

        protected override void InitCallbacks() {
        }

        string id2_Item;
        string id14_CustomUI;
        string id22_TypeLibPath;
        string id21_SourceItem;
        string id17_Pack;
        string id7_Language;
        string id10_DefaultImports;
        string id3_Name;
        string id6_CreateSandboxedAppDomain;
        string id16_Path;
        string id9_DefaultReferences;
        string id11_ExternalLibrary;
        string id12_Project;
        string id18_AssemblyPath;
        string id23_LoadFromBytes;
        string id13_Reference;
        string id5_ShadowCopyFiles;
        string id8_CompilerVersion;
        string id15_Image;
        string id19_ExplicitExports;
        string id1_DnaLibrary;
        string id4_RuntimeVersion;
        string id20_ComServer;

        protected override void InitIDs() {
            id2_Item = Reader.NameTable.Add(@"");
            id14_CustomUI = Reader.NameTable.Add(@"CustomUI");
            id22_TypeLibPath = Reader.NameTable.Add(@"TypeLibPath");
            id21_SourceItem = Reader.NameTable.Add(@"SourceItem");
            id17_Pack = Reader.NameTable.Add(@"Pack");
            id7_Language = Reader.NameTable.Add(@"Language");
            id10_DefaultImports = Reader.NameTable.Add(@"DefaultImports");
            id3_Name = Reader.NameTable.Add(@"Name");
            id6_CreateSandboxedAppDomain = Reader.NameTable.Add(@"CreateSandboxedAppDomain");
            id16_Path = Reader.NameTable.Add(@"Path");
            id9_DefaultReferences = Reader.NameTable.Add(@"DefaultReferences");
            id11_ExternalLibrary = Reader.NameTable.Add(@"ExternalLibrary");
            id12_Project = Reader.NameTable.Add(@"Project");
            id18_AssemblyPath = Reader.NameTable.Add(@"AssemblyPath");
            id23_LoadFromBytes = Reader.NameTable.Add(@"LoadFromBytes");
            id13_Reference = Reader.NameTable.Add(@"Reference");
            id5_ShadowCopyFiles = Reader.NameTable.Add(@"ShadowCopyFiles");
            id8_CompilerVersion = Reader.NameTable.Add(@"CompilerVersion");
            id15_Image = Reader.NameTable.Add(@"Image");
            id19_ExplicitExports = Reader.NameTable.Add(@"ExplicitExports");
            id1_DnaLibrary = Reader.NameTable.Add(@"DnaLibrary");
            id4_RuntimeVersion = Reader.NameTable.Add(@"RuntimeVersion");
            id20_ComServer = Reader.NameTable.Add(@"ComServer");
        }
    }

    public abstract class XmlSerializer1 : System.Xml.Serialization.XmlSerializer {
        protected override System.Xml.Serialization.XmlSerializationReader CreateReader() {
            return new XmlSerializationReaderDnaLibrary();
        }
        protected override System.Xml.Serialization.XmlSerializationWriter CreateWriter() {
            return new XmlSerializationWriterDnaLibrary();
        }
    }

    public sealed class DnaLibrarySerializer : XmlSerializer1 {

        public override System.Boolean CanDeserialize(System.Xml.XmlReader xmlReader) {
            return xmlReader.IsStartElement(@"DnaLibrary", @"");
        }

        protected override void Serialize(object objectToSerialize, System.Xml.Serialization.XmlSerializationWriter writer) {
            ((XmlSerializationWriterDnaLibrary)writer).Write8_DnaLibrary(objectToSerialize);
        }

        protected override object Deserialize(System.Xml.Serialization.XmlSerializationReader reader) {
            return ((XmlSerializationReaderDnaLibrary)reader).Read8_DnaLibrary();
        }
    }

    public class XmlSerializerContract : global::System.Xml.Serialization.XmlSerializerImplementation {
        public override global::System.Xml.Serialization.XmlSerializationReader Reader { get { return new XmlSerializationReaderDnaLibrary(); } }
        public override global::System.Xml.Serialization.XmlSerializationWriter Writer { get { return new XmlSerializationWriterDnaLibrary(); } }
        System.Collections.Hashtable readMethods = null;
        public override System.Collections.Hashtable ReadMethods {
            get {
                if (readMethods == null) {
                    System.Collections.Hashtable _tmp = new System.Collections.Hashtable();
                    _tmp[@"ExcelDna.Integration.DnaLibrary:::False:"] = @"Read8_DnaLibrary";
                    if (readMethods == null) readMethods = _tmp;
                }
                return readMethods;
            }
        }
        System.Collections.Hashtable writeMethods = null;
        public override System.Collections.Hashtable WriteMethods {
            get {
                if (writeMethods == null) {
                    System.Collections.Hashtable _tmp = new System.Collections.Hashtable();
                    _tmp[@"ExcelDna.Integration.DnaLibrary:::False:"] = @"Write8_DnaLibrary";
                    if (writeMethods == null) writeMethods = _tmp;
                }
                return writeMethods;
            }
        }
        System.Collections.Hashtable typedSerializers = null;
        public override System.Collections.Hashtable TypedSerializers {
            get {
                if (typedSerializers == null) {
                    System.Collections.Hashtable _tmp = new System.Collections.Hashtable();
                    _tmp.Add(@"ExcelDna.Integration.DnaLibrary:::False:", new DnaLibrarySerializer());
                    if (typedSerializers == null) typedSerializers = _tmp;
                }
                return typedSerializers;
            }
        }
        public override System.Boolean CanSerialize(System.Type type) {
            if (type == typeof(global::ExcelDna.Integration.DnaLibrary)) return true;
            return false;
        }
        public override System.Xml.Serialization.XmlSerializer GetSerializer(System.Type type) {
            if (type == typeof(global::ExcelDna.Integration.DnaLibrary)) return new DnaLibrarySerializer();
            return null;
        }
    }
}
