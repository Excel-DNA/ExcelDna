//#if _DYNAMIC_XMLSERIALIZER_COMPILATION
//[assembly:System.Security.AllowPartiallyTrustedCallers()]
//[assembly:System.Security.SecurityTransparent()]
//#endif
//[assembly:System.Reflection.AssemblyVersionAttribute("0.11.2629.36674")]
//[assembly:System.Xml.Serialization.XmlSerializerVersionAttribute(ParentAssemblyId=@"42f66455-e5a1-4e08-b8a9-564cdfa67981,", Version=@"2.0.0.0")]
namespace Microsoft.Xml.Serialization.GeneratedAssembly
{

    public class XmlSerializationWriterDnaLibrary : System.Xml.Serialization.XmlSerializationWriter
    {

        public void Write7_DnaLibrary(object o)
        {
            WriteStartDocument();
            if (o == null)
            {
                WriteEmptyTag(@"DnaLibrary", @"");
                return;
            }
            TopLevelElement();
            Write6_DnaLibrary(@"DnaLibrary", @"", ((global::ExcelDna.Integration.DnaLibrary)o), false, false);
        }

        void Write6_DnaLibrary(string n, string ns, global::ExcelDna.Integration.DnaLibrary o, bool isNullable, bool needType)
        {
            if ((object)o == null)
            {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType)
            {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.DnaLibrary))
                {
                }
                else
                {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"Language", @"", ((global::System.String)o.@Language));
            WriteAttribute(@"DefaultReferences", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultReferences)));
            WriteAttribute(@"DefaultImports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultImports)));
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.ExternalLibrary> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.ExternalLibrary>)o.@ExternalLibraries;
                if (a != null)
                {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++)
                    {
                        Write2_ExternalLibrary(@"ExternalLibrary", @"", ((global::ExcelDna.Integration.ExternalLibrary)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Project> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Project>)o.@Projects;
                if (a != null)
                {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++)
                    {
                        Write5_Project(@"Project", @"", ((global::ExcelDna.Integration.Project)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
                if (a != null)
                {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++)
                    {
                        Write3_Reference(@"Reference", @"", ((global::ExcelDna.Integration.Reference)a[ia]), false, false);
                    }
                }
            }
            {
                WriteValue(((global::System.String)o.@Code));
            }
            WriteEndElement(o);
        }

        void Write3_Reference(string n, string ns, global::ExcelDna.Integration.Reference o, bool isNullable, bool needType)
        {
            if ((object)o == null)
            {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType)
            {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.Reference))
                {
                }
                else
                {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"AssemblyPath", @"", ((global::System.String)o.@AssemblyPath));
            WriteEndElement(o);
        }

        void Write5_Project(string n, string ns, global::ExcelDna.Integration.Project o, bool isNullable, bool needType)
        {
            if ((object)o == null)
            {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType)
            {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.Project))
                {
                }
                else
                {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            WriteAttribute(@"Language", @"", ((global::System.String)o.@Language));
            WriteAttribute(@"DefaultReferences", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultReferences)));
            WriteAttribute(@"DefaultImports", @"", System.Xml.XmlConvert.ToString((global::System.Boolean)((global::System.Boolean)o.@DefaultImports)));
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
                if (a != null)
                {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++)
                    {
                        Write3_Reference(@"Reference", @"", ((global::ExcelDna.Integration.Reference)a[ia]), false, false);
                    }
                }
            }
            {
                global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem> a = (global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>)o.@SourceItems;
                if (a != null)
                {
                    for (int ia = 0; ia < ((System.Collections.ICollection)a).Count; ia++)
                    {
                        Write4_SourceItem(@"SourceItem", @"", ((global::ExcelDna.Integration.SourceItem)a[ia]), false, false);
                    }
                }
            }
            {
                WriteValue(((global::System.String)o.@Code));
            }
            WriteEndElement(o);
        }

        void Write4_SourceItem(string n, string ns, global::ExcelDna.Integration.SourceItem o, bool isNullable, bool needType)
        {
            if ((object)o == null)
            {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType)
            {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.SourceItem))
                {
                }
                else
                {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Name", @"", ((global::System.String)o.@Name));
            {
                WriteValue(((global::System.String)o.@Code));
            }
            WriteEndElement(o);
        }

        void Write2_ExternalLibrary(string n, string ns, global::ExcelDna.Integration.ExternalLibrary o, bool isNullable, bool needType)
        {
            if ((object)o == null)
            {
                if (isNullable) WriteNullTagLiteral(n, ns);
                return;
            }
            if (!needType)
            {
                System.Type t = o.GetType();
                if (t == typeof(global::ExcelDna.Integration.ExternalLibrary))
                {
                }
                else
                {
                    throw CreateUnknownTypeException(o);
                }
            }
            WriteStartElement(n, ns, o, false, null);
            if (needType) WriteXsiType(null, @"");
            WriteAttribute(@"Path", @"", ((global::System.String)o.@Path));
            WriteEndElement(o);
        }

        protected override void InitCallbacks()
        {
        }
    }

    public class XmlSerializationReaderDnaLibrary : System.Xml.Serialization.XmlSerializationReader
    {

        public object Read7_DnaLibrary()
        {
            object o = null;
            Reader.MoveToContent();
            if (Reader.NodeType == System.Xml.XmlNodeType.Element)
            {
                if (((object)Reader.LocalName == (object)id1_DnaLibrary && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o = Read6_DnaLibrary(false, true);
                }
                else
                {
                    throw CreateUnknownNodeException();
                }
            }
            else
            {
                UnknownNode(null, @":DnaLibrary");
            }
            return (object)o;
        }

        global::ExcelDna.Integration.DnaLibrary Read6_DnaLibrary(bool isNullable, bool checkType)
        {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType)
            {
                if (xsiType == null || ((object)((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object)((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item))
                {
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
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a_3 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
            bool[] paramsRead = new bool[8];
            while (Reader.MoveToNextAttribute())
            {
                if (!paramsRead[2] && ((object)Reader.LocalName == (object)id3_Name && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@Name = Reader.Value;
                    paramsRead[2] = true;
                }
                else if (!paramsRead[4] && ((object)Reader.LocalName == (object)id4_Language && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@Language = Reader.Value;
                    paramsRead[4] = true;
                }
                else if (!paramsRead[5] && ((object)Reader.LocalName == (object)id5_DefaultReferences && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[5] = true;
                }
                else if (!paramsRead[6] && ((object)Reader.LocalName == (object)id6_DefaultImports && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@DefaultImports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[6] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name))
                {
                    UnknownNode((object)o, @":Name, :Language, :DefaultReferences, :DefaultImports");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement)
            {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations0 = 0;
            int readerCount0 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None)
            {
                string tmp = null;
                if (Reader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    if (((object)Reader.LocalName == (object)id7_ExternalLibrary && (object)Reader.NamespaceURI == (object)id2_Item))
                    {
                        if ((object)(a_0) == null) Reader.Skip(); else a_0.Add(Read2_ExternalLibrary(false, true));
                    }
                    else if (((object)Reader.LocalName == (object)id8_Project && (object)Reader.NamespaceURI == (object)id2_Item))
                    {
                        if ((object)(a_1) == null) Reader.Skip(); else a_1.Add(Read5_Project(false, true));
                    }
                    else if (((object)Reader.LocalName == (object)id9_Reference && (object)Reader.NamespaceURI == (object)id2_Item))
                    {
                        if ((object)(a_3) == null) Reader.Skip(); else a_3.Add(Read3_Reference(false, true));
                    }
                    else
                    {
                        UnknownNode((object)o, @":ExternalLibrary, :Project, :Reference");
                    }
                }
                else if (Reader.NodeType == System.Xml.XmlNodeType.Text ||
                Reader.NodeType == System.Xml.XmlNodeType.CDATA ||
                Reader.NodeType == System.Xml.XmlNodeType.Whitespace ||
                Reader.NodeType == System.Xml.XmlNodeType.SignificantWhitespace)
                {
                    tmp = ReadString(tmp, false);
                    o.@Code = tmp;
                }
                else
                {
                    UnknownNode((object)o, @":ExternalLibrary, :Project, :Reference");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations0, ref readerCount0);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.Reference Read3_Reference(bool isNullable, bool checkType)
        {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType)
            {
                if (xsiType == null || ((object)((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object)((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item))
                {
                }
                else
                    throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.Reference o;
            o = new global::ExcelDna.Integration.Reference();
            bool[] paramsRead = new bool[1];
            while (Reader.MoveToNextAttribute())
            {
                if (!paramsRead[0] && ((object)Reader.LocalName == (object)id10_AssemblyPath && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@AssemblyPath = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name))
                {
                    UnknownNode((object)o, @":AssemblyPath");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement)
            {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations1 = 0;
            int readerCount1 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None)
            {
                if (Reader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    UnknownNode((object)o, @"");
                }
                else
                {
                    UnknownNode((object)o, @"");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations1, ref readerCount1);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.Project Read5_Project(bool isNullable, bool checkType)
        {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType)
            {
                if (xsiType == null || ((object)((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object)((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item))
                {
                }
                else
                    throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.Project o;
            try
            {
                o = (global::ExcelDna.Integration.Project)System.Activator.CreateInstance(typeof(global::ExcelDna.Integration.Project), System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.CreateInstance | System.Reflection.BindingFlags.NonPublic, null, new object[0], null);
            }
            catch (System.MissingMethodException)
            {
                throw CreateInaccessibleConstructorException(@"global::ExcelDna.Integration.Project");
            }
            catch (System.Security.SecurityException)
            {
                throw CreateCtorHasSecurityException(@"global::ExcelDna.Integration.Project");
            }
            if ((object)(o.@References) == null) o.@References = new global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference> a_2 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.Reference>)o.@References;
            if ((object)(o.@SourceItems) == null) o.@SourceItems = new global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>();
            global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem> a_5 = (global::System.Collections.Generic.List<global::ExcelDna.Integration.SourceItem>)o.@SourceItems;
            bool[] paramsRead = new bool[7];
            while (Reader.MoveToNextAttribute())
            {
                if (!paramsRead[0] && ((object)Reader.LocalName == (object)id3_Name && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!paramsRead[1] && ((object)Reader.LocalName == (object)id4_Language && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@Language = Reader.Value;
                    paramsRead[1] = true;
                }
                else if (!paramsRead[3] && ((object)Reader.LocalName == (object)id5_DefaultReferences && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@DefaultReferences = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[3] = true;
                }
                else if (!paramsRead[4] && ((object)Reader.LocalName == (object)id6_DefaultImports && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@DefaultImports = System.Xml.XmlConvert.ToBoolean(Reader.Value);
                    paramsRead[4] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name))
                {
                    UnknownNode((object)o, @":Name, :Language, :DefaultReferences, :DefaultImports");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement)
            {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations2 = 0;
            int readerCount2 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None)
            {
                string tmp = null;
                if (Reader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    if (((object)Reader.LocalName == (object)id9_Reference && (object)Reader.NamespaceURI == (object)id2_Item))
                    {
                        if ((object)(a_2) == null) Reader.Skip(); else a_2.Add(Read3_Reference(false, true));
                    }
                    else if (((object)Reader.LocalName == (object)id11_SourceItem && (object)Reader.NamespaceURI == (object)id2_Item))
                    {
                        if ((object)(a_5) == null) Reader.Skip(); else a_5.Add(Read4_SourceItem(false, true));
                    }
                    else
                    {
                        UnknownNode((object)o, @":Reference, :SourceItem");
                    }
                }
                else if (Reader.NodeType == System.Xml.XmlNodeType.Text ||
                Reader.NodeType == System.Xml.XmlNodeType.CDATA ||
                Reader.NodeType == System.Xml.XmlNodeType.Whitespace ||
                Reader.NodeType == System.Xml.XmlNodeType.SignificantWhitespace)
                {
                    tmp = ReadString(tmp, false);
                    o.@Code = tmp;
                }
                else
                {
                    UnknownNode((object)o, @":Reference, :SourceItem");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations2, ref readerCount2);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.SourceItem Read4_SourceItem(bool isNullable, bool checkType)
        {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType)
            {
                if (xsiType == null || ((object)((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object)((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item))
                {
                }
                else
                    throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.SourceItem o;
            o = new global::ExcelDna.Integration.SourceItem();
            bool[] paramsRead = new bool[2];
            while (Reader.MoveToNextAttribute())
            {
                if (!paramsRead[0] && ((object)Reader.LocalName == (object)id3_Name && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@Name = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name))
                {
                    UnknownNode((object)o, @":Name");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement)
            {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations3 = 0;
            int readerCount3 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None)
            {
                string tmp = null;
                if (Reader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    UnknownNode((object)o, @"");
                }
                else if (Reader.NodeType == System.Xml.XmlNodeType.Text ||
                Reader.NodeType == System.Xml.XmlNodeType.CDATA ||
                Reader.NodeType == System.Xml.XmlNodeType.Whitespace ||
                Reader.NodeType == System.Xml.XmlNodeType.SignificantWhitespace)
                {
                    tmp = ReadString(tmp, false);
                    o.@Code = tmp;
                }
                else
                {
                    UnknownNode((object)o, @"");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations3, ref readerCount3);
            }
            ReadEndElement();
            return o;
        }

        global::ExcelDna.Integration.ExternalLibrary Read2_ExternalLibrary(bool isNullable, bool checkType)
        {
            System.Xml.XmlQualifiedName xsiType = checkType ? GetXsiType() : null;
            bool isNull = false;
            if (isNullable) isNull = ReadNull();
            if (checkType)
            {
                if (xsiType == null || ((object)((System.Xml.XmlQualifiedName)xsiType).Name == (object)id2_Item && (object)((System.Xml.XmlQualifiedName)xsiType).Namespace == (object)id2_Item))
                {
                }
                else
                    throw CreateUnknownTypeException((System.Xml.XmlQualifiedName)xsiType);
            }
            if (isNull) return null;
            global::ExcelDna.Integration.ExternalLibrary o;
            o = new global::ExcelDna.Integration.ExternalLibrary();
            bool[] paramsRead = new bool[1];
            while (Reader.MoveToNextAttribute())
            {
                if (!paramsRead[0] && ((object)Reader.LocalName == (object)id12_Path && (object)Reader.NamespaceURI == (object)id2_Item))
                {
                    o.@Path = Reader.Value;
                    paramsRead[0] = true;
                }
                else if (!IsXmlnsAttribute(Reader.Name))
                {
                    UnknownNode((object)o, @":Path");
                }
            }
            Reader.MoveToElement();
            if (Reader.IsEmptyElement)
            {
                Reader.Skip();
                return o;
            }
            Reader.ReadStartElement();
            Reader.MoveToContent();
            int whileIterations4 = 0;
            int readerCount4 = ReaderCount;
            while (Reader.NodeType != System.Xml.XmlNodeType.EndElement && Reader.NodeType != System.Xml.XmlNodeType.None)
            {
                if (Reader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    UnknownNode((object)o, @"");
                }
                else
                {
                    UnknownNode((object)o, @"");
                }
                Reader.MoveToContent();
                CheckReaderCount(ref whileIterations4, ref readerCount4);
            }
            ReadEndElement();
            return o;
        }

        protected override void InitCallbacks()
        {
        }

        string id11_SourceItem;
        string id3_Name;
        string id8_Project;
        string id6_DefaultImports;
        string id12_Path;
        string id10_AssemblyPath;
        string id2_Item;
        string id5_DefaultReferences;
        string id4_Language;
        string id7_ExternalLibrary;
        string id1_DnaLibrary;
        string id9_Reference;

        protected override void InitIDs()
        {
            id11_SourceItem = Reader.NameTable.Add(@"SourceItem");
            id3_Name = Reader.NameTable.Add(@"Name");
            id8_Project = Reader.NameTable.Add(@"Project");
            id6_DefaultImports = Reader.NameTable.Add(@"DefaultImports");
            id12_Path = Reader.NameTable.Add(@"Path");
            id10_AssemblyPath = Reader.NameTable.Add(@"AssemblyPath");
            id2_Item = Reader.NameTable.Add(@"");
            id5_DefaultReferences = Reader.NameTable.Add(@"DefaultReferences");
            id4_Language = Reader.NameTable.Add(@"Language");
            id7_ExternalLibrary = Reader.NameTable.Add(@"ExternalLibrary");
            id1_DnaLibrary = Reader.NameTable.Add(@"DnaLibrary");
            id9_Reference = Reader.NameTable.Add(@"Reference");
        }
    }

    public abstract class XmlSerializer1 : System.Xml.Serialization.XmlSerializer
    {
        protected override System.Xml.Serialization.XmlSerializationReader CreateReader()
        {
            return new XmlSerializationReaderDnaLibrary();
        }
        protected override System.Xml.Serialization.XmlSerializationWriter CreateWriter()
        {
            return new XmlSerializationWriterDnaLibrary();
        }
    }

    public sealed class DnaLibrarySerializer : XmlSerializer1
    {

        public override System.Boolean CanDeserialize(System.Xml.XmlReader xmlReader)
        {
            return xmlReader.IsStartElement(@"DnaLibrary", @"");
        }

        protected override void Serialize(object objectToSerialize, System.Xml.Serialization.XmlSerializationWriter writer)
        {
            ((XmlSerializationWriterDnaLibrary)writer).Write7_DnaLibrary(objectToSerialize);
        }

        protected override object Deserialize(System.Xml.Serialization.XmlSerializationReader reader)
        {
            return ((XmlSerializationReaderDnaLibrary)reader).Read7_DnaLibrary();
        }
    }

    public class XmlSerializerContract : global::System.Xml.Serialization.XmlSerializerImplementation
    {
        public override global::System.Xml.Serialization.XmlSerializationReader Reader { get { return new XmlSerializationReaderDnaLibrary(); } }
        public override global::System.Xml.Serialization.XmlSerializationWriter Writer { get { return new XmlSerializationWriterDnaLibrary(); } }
        System.Collections.Hashtable readMethods = null;
        public override System.Collections.Hashtable ReadMethods
        {
            get
            {
                if (readMethods == null)
                {
                    System.Collections.Hashtable _tmp = new System.Collections.Hashtable();
                    _tmp[@"ExcelDna.Integration.DnaLibrary:::False:"] = @"Read7_DnaLibrary";
                    if (readMethods == null) readMethods = _tmp;
                }
                return readMethods;
            }
        }
        System.Collections.Hashtable writeMethods = null;
        public override System.Collections.Hashtable WriteMethods
        {
            get
            {
                if (writeMethods == null)
                {
                    System.Collections.Hashtable _tmp = new System.Collections.Hashtable();
                    _tmp[@"ExcelDna.Integration.DnaLibrary:::False:"] = @"Write7_DnaLibrary";
                    if (writeMethods == null) writeMethods = _tmp;
                }
                return writeMethods;
            }
        }
        System.Collections.Hashtable typedSerializers = null;
        public override System.Collections.Hashtable TypedSerializers
        {
            get
            {
                if (typedSerializers == null)
                {
                    System.Collections.Hashtable _tmp = new System.Collections.Hashtable();
                    _tmp.Add(@"ExcelDna.Integration.DnaLibrary:::False:", new DnaLibrarySerializer());
                    if (typedSerializers == null) typedSerializers = _tmp;
                }
                return typedSerializers;
            }
        }
        public override System.Boolean CanSerialize(System.Type type)
        {
            if (type == typeof(global::ExcelDna.Integration.DnaLibrary)) return true;
            return false;
        }
        public override System.Xml.Serialization.XmlSerializer GetSerializer(System.Type type)
        {
            if (type == typeof(global::ExcelDna.Integration.DnaLibrary)) return new DnaLibrarySerializer();
            return null;
        }
    }
}
