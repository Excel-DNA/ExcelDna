//#if _DYNAMIC_XMLSERIALIZER_COMPILATION
//[assembly:System.Security.AllowPartiallyTrustedCallers()]
//[assembly:System.Security.SecurityTransparent()]
//[assembly:System.Security.SecurityRules(System.Security.SecurityRuleSet.Level1)]
//#endif
//[assembly:System.Reflection.AssemblyVersionAttribute("0.31.5019.3238")]
//[assembly:System.Xml.Serialization.XmlSerializerVersionAttribute(ParentAssemblyId=@"cf16aa56-f14f-463f-ab55-a4d6d6ddb3be,", Version=@"4.0.0.0")]
namespace ExcelDna.Serialization
{
    public class XmlSerializationWriterDnaLibrary : System.Xml.Serialization.XmlSerializationWriter {

        public void Write8_DnaLibrary(object o) {
            WriteStartDocument();
            if (o == null) {
                WriteEmptyTag(@"DnaLibrary", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
                return;
            }
            TopLevelElement();
            Write7_DnaLibrary(@"DnaLibrary", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", ((global::ExcelDna.Integration.DnaLibrary)o), false, false);
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
            if (needType) WriteXsiType(null, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
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
                        Write2_ExternalLibrary(@"ExternalLibrary", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", ((global::ExcelDna.Integration.ExternalLibrary)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Project> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Project>)o.@Projects;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write5_Project(@"Project", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", ((global::ExcelDna.Integration.Project)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write3_Reference(@"Reference", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", ((global::ExcelDna.Integration.Reference)a[ia]), false, false);
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
                            WriteElementLiteral((System.Xml.XmlNode)((global::System.Xml.XmlNode)a[ia]), @"CustomUI", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", false, false);
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
                        Write6_Image(@"Image", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", ((global::ExcelDna.Integration.Image)a[ia]), false, false);
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
            if (needType) WriteXsiType(null, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
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
            if (needType) WriteXsiType(null, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"AssemblyPath", @"", ((global::System.String)o.@AssemblyPath));
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
            WriteAttribute(@"IncludePdb", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@IncludePdb)));
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
            if (needType) WriteXsiType(null, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"Language", @"", ((global::System.String)o.@Language));
            WriteAttribute(@"CompilerVersion", @"", ((global::System.String)o.@CompilerVersion));
            WriteAttribute(@"DefaultReferences", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultReferences)));
            WriteAttribute(@"DefaultImports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultImports)));
            WriteAttribute(@"ExplicitExports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ExplicitExports)));
            WriteAttribute(@"ExplicitRegistration", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ExplicitRegistration)));
            WriteAttribute(@"ComServer", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ComServer)));
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write3_Reference(@"Reference", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", ((global::ExcelDna.Integration.Reference)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>)o.@SourceItems;
                if (a != null) {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++) {
                        Write4_SourceItem(@"SourceItem", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary", ((global::ExcelDna.Integration.SourceItem)a[ia]), false, false);
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
            if (needType) WriteXsiType(null, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
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
            if (needType) WriteXsiType(null, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
            WriteAttribute(@"Path", @"", ((global::System.String)o.@Path));
            WriteAttribute(@"TypeLibPath", @"", ((global::System.String)o.@TypeLibPath));
            WriteAttribute(@"ComServer", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ComServer)));
            WriteAttribute(@"Pack", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@Pack)));
            WriteAttribute(@"LoadFromBytes", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@LoadFromBytes)));
            WriteAttribute(@"ExplicitExports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ExplicitExports)));
            WriteAttribute(@"ExplicitRegistration", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@ExplicitRegistration)));
            WriteAttribute(@"UseVersionAsOutputVersion", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@UseVersionAsOutputVersion)));
            WriteAttribute(@"IncludePdb", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@IncludePdb)));
            WriteEndElement(o);
        }

        protected override void InitCallbacks() {
        }
    }

    public class XmlSerializationReaderDnaLibrary : System.Xml.Serialization.XmlSerializationReader {
        /*
          ### IMPORTANT
          This class is auto-generated via sgen, but have been manually edited
          Make sure you apply the same edits if you re-generate the code again

          More details in PR #194
          https://github.com/Excel-DNA/ExcelDna/pull/194

          Manual edit scope:
            **All checks and references to id2_Item (which translates to xmlns) have been commented out**

          This is the command line that we previously used to generate this file:
          "sgen /a:bin\Debug\exceldna.integration.dll /t:ExcelDna.Integration.DnaLibrary /k /f /out:."
        */

        public object Read8_DnaLibrary() {
            object o = null;
            Reader.MoveToContent();
            if (Reader.NodeType == System.Xml.XmlNodeType.Element) {
                if (((object) Reader.LocalName == (object)id1_DnaLibrary /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                    o = Read7_DnaLibrary(false, true);
                }
                else {
                    throw CreateUnknownNodeException();
                }
            }
            else {
                UnknownNode(null, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary:DnaLibrary");
            }
            return (object)o;
        }

        global::ExcelDna.Integration.DnaLibrary Read7_DnaLibrary(bool isNullable, bool checkType) {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType) {
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id3_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
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
                if (!paramsRead[2] && ((object) Reader.LocalName == (object)id4_Name && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[2] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id5_RuntimeVersion && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@RuntimeVersion = Reader.Value;
                    paramsRead[3] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id6_ShadowCopyFiles && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@ShadowCopyFiles = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id7_CreateSandboxedAppDomain && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@CreateSandboxedAppDomain = Reader.Value;
                    paramsRead[5] = true;
                }
                else if (!paramsRead[7] && ((object) Reader.LocalName == (object)id8_Language && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Language = Reader.Value;
                    paramsRead[7] = true;
                }
                else if (!paramsRead[8] && ((object) Reader.LocalName == (object)id9_CompilerVersion && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@CompilerVersion = Reader.Value;
                    paramsRead[8] = true;
                }
                else if (!paramsRead[9] && ((object) Reader.LocalName == (object)id10_DefaultReferences && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[9] = true;
                }
                else if (!paramsRead[10] && ((object) Reader.LocalName == (object)id11_DefaultImports && (object) Reader.NamespaceURI == (object)id3_Item)) {
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
                    if (((object) Reader.LocalName == (object)id12_ExternalLibrary /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                        if ((object)(a_0) == null) Reader.Skip(); else a_0.Add(Read2_ExternalLibrary(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id13_Project /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                        if ((object)(a_1) == null) Reader.Skip(); else a_1.Add(Read5_Project(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id14_Reference /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                        if ((object)(a_6) == null) Reader.Skip(); else a_6.Add(Read3_Reference(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id15_CustomUI /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                        a_12.Add((global::System.Xml.XmlNode)ReadXmlNode(true));
                    }
                    else if (((object) Reader.LocalName == (object)id16_Image /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                        if ((object)(a_13) == null) Reader.Skip(); else a_13.Add(Read6_Image(false, true));
                    }
                    else {
                        UnknownNode((object)o, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary:ExternalLibrary, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Project, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Reference, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:CustomUI, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Image");
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
                    UnknownNode((object)o, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary:ExternalLibrary, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Project, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Reference, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:CustomUI, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Image");
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
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id3_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.Image o;
            o = new global::ExcelDna.Integration.Image();
            bool[] paramsRead = new bool[3];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id4_Name && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id17_Path && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id18_Pack && (object) Reader.NamespaceURI == (object)id3_Item)) {
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
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id3_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.Reference o;
            o = new global::ExcelDna.Integration.Reference();
            bool[] paramsRead = new bool[5];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id4_Name && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id19_AssemblyPath && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@AssemblyPath = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id18_Pack && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[2] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id20_IncludePdb && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@IncludePdb = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[3] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id17_Path && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[4] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :AssemblyPath, :Pack, :IncludePdb, :Path");
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
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id3_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
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
            global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem> a_9 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>)o.@SourceItems;
            bool[] paramsRead = new bool[11];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id4_Name && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id8_Language && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Language = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id9_CompilerVersion && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@CompilerVersion = Reader.Value;
                    paramsRead[2] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id10_DefaultReferences && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id11_DefaultImports && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@DefaultImports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[5] = true;
                }
                else if (!paramsRead[6] && ((object) Reader.LocalName == (object)id21_ExplicitExports && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@ExplicitExports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[6] = true;
                }
                else if (!paramsRead[7] && ((object) Reader.LocalName == (object)id22_ExplicitRegistration && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@ExplicitRegistration = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[7] = true;
                }
                else if (!paramsRead[8] && ((object) Reader.LocalName == (object)id23_ComServer && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@ComServer = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[8] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Name, :Language, :CompilerVersion, :DefaultReferences, :DefaultImports, :ExplicitExports, :ExplicitRegistration, :ComServer");
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
                    if (((object) Reader.LocalName == (object)id14_Reference /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                        if ((object)(a_3) == null) Reader.Skip(); else a_3.Add(Read3_Reference(false, true));
                    }
                    else if (((object) Reader.LocalName == (object)id24_SourceItem /*&& (object) Reader.NamespaceURI == (object)id2_Item*/)) {
                        if ((object)(a_9) == null) Reader.Skip(); else a_9.Add(Read4_SourceItem(false, true));
                    }
                    else {
                        UnknownNode((object)o, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Reference, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:SourceItem");
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
                    UnknownNode((object)o, @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary:Reference, http://schemas.excel-dna.net/addin/2020/07/dnalibrary:SourceItem");
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
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id3_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.SourceItem o;
            o = new global::ExcelDna.Integration.SourceItem();
            bool[] paramsRead = new bool[4];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id18_Pack && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id4_Name && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Name = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id17_Path && (object) Reader.NamespaceURI == (object)id3_Item)) {
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
            if (xsiType == null || ((object) ((System.Xml.XmlQualifiedName)xsiType).Name == (object)id3_Item && (object) ((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item)) {
            }
            else
                throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.ExternalLibrary o;
            o = new global::ExcelDna.Integration.ExternalLibrary();
            bool[] paramsRead = new bool[9];
            while (Reader.MoveToNextAttribute()) {
                if (!paramsRead[0] && ((object) Reader.LocalName == (object)id17_Path && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Path = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object) Reader.LocalName == (object)id25_TypeLibPath && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@TypeLibPath = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[2] && ((object) Reader.LocalName == (object)id23_ComServer && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@ComServer = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[2] = true;
                }
                else if (!paramsRead[3] && ((object) Reader.LocalName == (object)id18_Pack && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@Pack = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[3] = true;
                }
                else if (!paramsRead[4] && ((object) Reader.LocalName == (object)id26_LoadFromBytes && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@LoadFromBytes = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object) Reader.LocalName == (object)id21_ExplicitExports && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@ExplicitExports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[5] = true;
                }
                else if (!paramsRead[6] && ((object) Reader.LocalName == (object)id22_ExplicitRegistration && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@ExplicitRegistration = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[6] = true;
                }
                else if (!paramsRead[7] && ((object) Reader.LocalName == (object)id27_UseVersionAsOutputVersion && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@UseVersionAsOutputVersion = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[7] = true;
                }
                else if (!paramsRead[8] && ((object) Reader.LocalName == (object)id20_IncludePdb && (object) Reader.NamespaceURI == (object)id3_Item)) {
                    o.@IncludePdb = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[8] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name)) {
                    UnknownNode((object)o, @":Path, :TypeLibPath, :ComServer, :Pack, :LoadFromBytes, :ExplicitExports, :ExplicitRegistration, :UseVersionAsOutputVersion, :IncludePdb");
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

        string id22_ExplicitRegistration;
        string id18_Pack;
        string id24_SourceItem;
        string id5_RuntimeVersion;
        string id8_Language;
        string id20_IncludePdb;
        string id16_Image;
        string id27_UseVersionAsOutputVersion;
        string id11_DefaultImports;
        string id2_Item;
        string id17_Path;
        string id10_DefaultReferences;
        string id6_ShadowCopyFiles;
        string id7_CreateSandboxedAppDomain;
        string id21_ExplicitExports;
        string id19_AssemblyPath;
        string id14_Reference;
        string id26_LoadFromBytes;
        string id12_ExternalLibrary;
        string id23_ComServer;
        string id25_TypeLibPath;
        string id15_CustomUI;
        string id1_DnaLibrary;
        string id4_Name;
        string id9_CompilerVersion;
        string id3_Item;
        string id13_Project;

        protected override void InitIDs() {
            id22_ExplicitRegistration = Reader.NameTable.Add(@"ExplicitRegistration");
            id18_Pack = Reader.NameTable.Add(@"Pack");
            id24_SourceItem = Reader.NameTable.Add(@"SourceItem");
            id5_RuntimeVersion = Reader.NameTable.Add(@"RuntimeVersion");
            id8_Language = Reader.NameTable.Add(@"Language");
            id20_IncludePdb = Reader.NameTable.Add(@"IncludePdb");
            id16_Image = Reader.NameTable.Add(@"Image");
            id27_UseVersionAsOutputVersion = Reader.NameTable.Add(@"UseVersionAsOutputVersion");
            id11_DefaultImports = Reader.NameTable.Add(@"DefaultImports");
            id2_Item = Reader.NameTable.Add(@"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
            id17_Path = Reader.NameTable.Add(@"Path");
            id10_DefaultReferences = Reader.NameTable.Add(@"DefaultReferences");
            id6_ShadowCopyFiles = Reader.NameTable.Add(@"ShadowCopyFiles");
            id7_CreateSandboxedAppDomain = Reader.NameTable.Add(@"CreateSandboxedAppDomain");
            id21_ExplicitExports = Reader.NameTable.Add(@"ExplicitExports");
            id19_AssemblyPath = Reader.NameTable.Add(@"AssemblyPath");
            id14_Reference = Reader.NameTable.Add(@"Reference");
            id26_LoadFromBytes = Reader.NameTable.Add(@"LoadFromBytes");
            id12_ExternalLibrary = Reader.NameTable.Add(@"ExternalLibrary");
            id23_ComServer = Reader.NameTable.Add(@"ComServer");
            id25_TypeLibPath = Reader.NameTable.Add(@"TypeLibPath");
            id15_CustomUI = Reader.NameTable.Add(@"CustomUI");
            id1_DnaLibrary = Reader.NameTable.Add(@"DnaLibrary");
            id4_Name = Reader.NameTable.Add(@"Name");
            id9_CompilerVersion = Reader.NameTable.Add(@"CompilerVersion");
            id3_Item = Reader.NameTable.Add(@"");
            id13_Project = Reader.NameTable.Add(@"Project");
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
            return xmlReader.IsStartElement(@"DnaLibrary", @"http://schemas.excel-dna.net/addin/2020/07/dnalibrary");
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
                    _tmp[@"ExcelDna.Integration.DnaLibrary:http://schemas.excel-dna.net/addin/2020/07/dnalibrary::False:"] = @"Read8_DnaLibrary";
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
                    _tmp[@"ExcelDna.Integration.DnaLibrary:http://schemas.excel-dna.net/addin/2020/07/dnalibrary::False:"] = @"Write8_DnaLibrary";
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
                    _tmp.Add(@"ExcelDna.Integration.DnaLibrary:http://schemas.excel-dna.net/addin/2020/07/dnalibrary::False:", new DnaLibrarySerializer());
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
