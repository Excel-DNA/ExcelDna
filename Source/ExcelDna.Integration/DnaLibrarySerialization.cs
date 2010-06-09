//#if _DYNAMIC_XMLSERIALIZER_COMPILATION
//[assembly:System.Security.AllowPartiallyTrustedCallers()]
//[assembly:System.Security.SecurityTransparent()]
//#endif
//[assembly:System.Reflection.AssemblyVersionAttribute("0.26.3792.1875")]
//[assembly:System.Xml.Serialization.XmlSerializerVersionAttribute(ParentAssemblyId=@"d90fa6e2-81f0-4242-89ac-cde68a8074b4,", Version=@"2.0.0.0")]
namespace Microsoft.Xml.Serialization.GeneratedAssembly {

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
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Image> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Image>)o.Images;
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
            WriteAttribute(@"AssemblyPath", @"", ((global::System.String)o.@AssemblyPath));
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
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
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
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
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
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
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a_4 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
            if ((object)(o.@CustomUIs) == null) o.@CustomUIs = new global::System.Collections.Generic.List<global::System.Xml.XmlNode>();
            global::System.Collections.Generic.List<global::System.Xml.XmlNode> a_10 = (global::System.Collections.Generic.List<global::System.Xml.XmlNode>)o.@CustomUIs;
            if ((object)(o.Images) == null) o.Images = new global::System.Collections.Generic.List<global::ExcelDna.Integration.Image>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Image> a_11 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Image>)o.Images;
            bool[] paramsRead = new bool[12];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[2] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[2] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id4_RuntimeVersion && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@RuntimeVersion = Reader.Value;
                    paramsRead[3] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id5_Language && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Language = Reader.Value;
                    paramsRead[5] = true;
                }
                else if (!paramsRead[6] && ((object) Reader.LocalName == (object)id6_CompilerVersion && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@CompilerVersion = Reader.Value;
                    paramsRead[6] = true;
                }
                else if (!paramsRead[7] && ((object) Reader.LocalName == (object)id7_DefaultReferences && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[7] = true;
                }
                else if (!paramsRead[8] && ((object) Reader.LocalName == (object)id8_DefaultImports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultImports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[8] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :RuntimeVersion, :Language, :CompilerVersion, :DefaultReferences, :DefaultImports");
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
                    if (((object) Reader.LocalName == (object)id9_ExternalLibrary && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_0) == null) Reader.Skip(); else a_0.Add(Read2_ExternalLibrary(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id10_Project && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_1) == null) Reader.Skip(); else a_1.Add(Read5_Project(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id11_Reference && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_4) == null) Reader.Skip(); else a_4.Add(Read3_Reference(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id12_CustomUI && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        a_10.Add((global::System.Xml.XmlNode)ReadXmlNode(true));
                    }
                    else if (((object) Reader.LocalName == (object)id13_Image && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_11) == null) Reader.Skip(); else a_11.Add(Read6_Image(false, true));
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
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id14_Path && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id15_Pack && (object) Reader.NamespaceURI == (object)id2_Item)) {
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
            bool[] paramsRead = new bool[3];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id16_AssemblyPath && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@AssemblyPath = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id15_Pack && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[2] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":AssemblyPath, :Name, :Pack");
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
            global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem> a_7 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>)o.@SourceItems;
            bool[] paramsRead = new bool[9];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id5_Language && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Language = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id6_CompilerVersion && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@CompilerVersion = Reader.Value;
                    paramsRead[2] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id7_DefaultReferences && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id8_DefaultImports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@DefaultImports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[5] = true;
                }
                else if (!paramsRead[6] && ((object) Reader.LocalName == (object)id17_ExplicitExports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@ExplicitExports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[6] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :Language, :CompilerVersion, :DefaultReferences, :DefaultImports, :ExplicitExports");
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
                    if (((object) Reader.LocalName == (object)id11_Reference && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_3) == null) Reader.Skip(); else a_3.Add(Read3_Reference(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id18_SourceItem && (object) Reader.NamespaceURI == (object)id2_Item)) {
                        if ((object)(a_7) == null) Reader.Skip(); else a_7.Add(Read4_SourceItem(false, true));
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
            bool[] paramsRead = new bool[2];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id3_Name && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name");
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
            bool[] paramsRead = new bool[3];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id14_Path && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id15_Pack && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id17_ExplicitExports && (object) Reader.NamespaceURI == (object)id2_Item)) {
                    o.@ExplicitExports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[2] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Path, :Pack, :ExplicitExports");
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

        string id7_DefaultReferences;
        string id3_Name;
        string id16_AssemblyPath;
        string id10_Project;
        string id14_Path;
        string id13_Image;
        string id6_CompilerVersion;
        string id18_SourceItem;
        string id1_DnaLibrary;
        string id12_CustomUI;
        string id2_Item;
        string id15_Pack;
        string id11_Reference;
        string id8_DefaultImports;
        string id5_Language;
        string id4_RuntimeVersion;
        string id17_ExplicitExports;
        string id9_ExternalLibrary;

        protected override void InitIDs() {
            id7_DefaultReferences = Reader.NameTable.Add(@"DefaultReferences");
            id3_Name = Reader.NameTable.Add(@"Name");
            id16_AssemblyPath = Reader.NameTable.Add(@"AssemblyPath");
            id10_Project = Reader.NameTable.Add(@"Project");
            id14_Path = Reader.NameTable.Add(@"Path");
            id13_Image = Reader.NameTable.Add(@"Image");
            id6_CompilerVersion = Reader.NameTable.Add(@"CompilerVersion");
            id18_SourceItem = Reader.NameTable.Add(@"SourceItem");
            id1_DnaLibrary = Reader.NameTable.Add(@"DnaLibrary");
            id12_CustomUI = Reader.NameTable.Add(@"CustomUI");
            id2_Item = Reader.NameTable.Add(@"");
            id15_Pack = Reader.NameTable.Add(@"Pack");
            id11_Reference = Reader.NameTable.Add(@"Reference");
            id8_DefaultImports = Reader.NameTable.Add(@"DefaultImports");
            id5_Language = Reader.NameTable.Add(@"Language");
            id4_RuntimeVersion = Reader.NameTable.Add(@"RuntimeVersion");
            id17_ExplicitExports = Reader.NameTable.Add(@"ExplicitExports");
            id9_ExternalLibrary = Reader.NameTable.Add(@"ExternalLibrary");
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
