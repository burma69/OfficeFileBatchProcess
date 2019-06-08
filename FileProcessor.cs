using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ionic.Zip;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
using System.Xml.XPath;
using System.Security;

namespace OfficeBatchProcess
{
    class FileProcessor : IDisposable
    {
        public int ProcessFile(string SourceFileName, string DestFileName)
        {
            return ProcessFile(File.ReadAllBytes(SourceFileName), SourceFileName, DestFileName);
        }

        public int ProcessFile(byte[] SourceFile, string SourceFileName, string DestFileName)
        {
            int ChangesCount = 0;
            Console.WriteLine();
            MemoryStream ms = new MemoryStream(SourceFile, false);

            if (System.IO.File.Exists(DestFileName))
            {
                Console.WriteLine("File already exists. Skipping [" + DestFileName + "]");
                return ChangesCount;
            }

            var options = new ReadOptions ();//{ StatusMessageWriter = System.Console.Out };
            using (ZipFile zipSource = ZipFile.Read(ms, options))
            {
                using (ZipFile zipDest = new ZipFile())
                {
                    foreach (var entry in zipSource)
                    {
                        MemoryStream dest = new MemoryStream();
                       
                        ChangesCount += ProcessEntry(entry, zipDest, SourceFileName, dest, DestFileName+" - extracts.xml");
                        ZipEntry d = zipDest.AddEntry(entry.FileName, dest);                        
                    }

                    //if (ChangesCount == 0) System.IO.File.Copy(SourceFileName, DestFileName, true);
                    //else 
                    if (ChangesCount > 0)
                    {
                        Console.WriteLine("Saving [" + DestFileName + "]");
                        zipDest.ParallelDeflateThreshold = -1; //memory leak otherwise
                        Directory.CreateDirectory(Path.GetDirectoryName(DestFileName));
                        zipDest.Save(DestFileName);                        
                    }
                }                
            }

            return ChangesCount;
        }

        private int CascadeRemove(XmlNode ParentNode, XmlNamespaceManager nsmgr, int ParentKey,
            BatchConfigSection.ProcessParamsElementCollection.ParamElement.DeduplicateElementsElementCollection.DeduplicateElement.CascadeRemovesElementCollection crc,
            int CurrentCascade, int StackLimit = 5)
        {
            
            int ChangesCount = 0;
            BatchConfigSection.ProcessParamsElementCollection.ParamElement.DeduplicateElementsElementCollection.DeduplicateElement.CascadeRemovesElementCollection.CascadeRemoveElement cr;

            if (StackLimit < 0) return ChangesCount;         
            cr = crc[CurrentCascade];
            


            if (cr.UpdateLinkedIndexElement != "")
            {
                string FilterXpath = cr.UpdateLinkedIndexElement;
                if (cr.ParentAttributeNameForXPathDash != "")
                    FilterXpath = FilterXpath.Replace("#", ParentNode.Attributes[cr.ParentAttributeNameForXPathDash].Value);

                XmlNodeList LinkedNodes = ParentNode.SelectNodes(FilterXpath, nsmgr);

                foreach(XmlNode l in LinkedNodes) 
                {
                    int j = int.Parse(l.Attributes[cr.UpdateLinkedIndexAttribute].Value);
                    if (j > ParentKey)
                    {
                        l.Attributes[cr.UpdateLinkedIndexAttribute].Value = (j - 1).ToString();
                        ChangesCount += 1;
                    }
                    else if (j == ParentKey)
                    {
                        int ThisKey = -1;

                        for (int w = 0; w < l.ParentNode.ChildNodes.Count; w++)
                            if (l.ParentNode.ChildNodes[w] == l)
                                ThisKey = w;

                        //((XmlElement)(l)).SetAttribute("DEBUG", "DELETE"+(6-StackLimit).ToString()+"-"+ThisKey.ToString()+"-"+ParentKey.ToString());
                        l.ParentNode.RemoveChild(l);
                        ChangesCount += 1;

                        for (int c = 0; c < crc.Count; c++)
                            if (crc[c].CascadeFrom == cr.Id)
                                ChangesCount += CascadeRemove(l, nsmgr, ThisKey, crc, c, StackLimit - 1);
                    }
                }
            }

            return ChangesCount;
        }

        private void ExtractContents(StreamWriter ExtractWriter, XmlNode n, XmlNamespaceManager nsmgr, string SliceByXPaths)
        {
            ExtractWriter.Write("<" + n.Name + ">");
            if (SliceByXPaths == "") ExtractWriter.Write(n.InnerText);
            else
            {
                string[] XPathNodes = SliceByXPaths.Split(',');

                XmlNodeList nodeList = n.SelectNodes(XPathNodes[0], nsmgr);
                foreach (XmlNode innerNode in nodeList)
                {
                    if (XPathNodes.Length > 1)
                    ExtractContents(ExtractWriter, innerNode, nsmgr, SliceByXPaths.Remove(0, XPathNodes[0].Length + 1));
                    else ExtractWriter.Write(SecurityElement.Escape(innerNode.InnerText));

                }
            }

            ExtractWriter.Write("</" + n.Name + ">");
        }

        public int ProcessEntry(ZipEntry e, ZipFile zipDest, string SourceFileName, MemoryStream dest, string ExtractsFileName)
        {
            int ChangesCount = 0;
            XmlDocument XMLDoc = null;
            StreamWriter ExtractWriter = null;
            

            foreach (BatchConfigSection.ProcessParamsElementCollection.ParamElement p in BatchConfig.Config.ProcessParams)
            {
                if (Regex.IsMatch(SourceFileName, p.XLSName, RegexOptions.IgnoreCase))
                    if (Regex.IsMatch(e.FileName, p.ZipObjectPath, RegexOptions.IgnoreCase))
                    {
                        Console.WriteLine("[" + e.FileName + "] rule " + p.XLSName + ", " + p.ZipObjectPath);

                        if (XMLDoc == null) 
                        using (MemoryStream source = new MemoryStream((int)e.UncompressedSize))
                        {                            
	                        e.Extract(source);                        
                            XMLDoc = new XmlDocument();
                            source.Position = 0;
                            XMLDoc.Load(source);
                        }

                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(XMLDoc.NameTable);
                        nsmgr.AddNamespace(p.NodeXPath[1].ToString(), p.NodeNamespace);
                        
                        XmlNodeList nodeList = XMLDoc.DocumentElement.SelectNodes(p.NodeXPath, nsmgr);

                        foreach (XmlNode n in nodeList)
                        {
                            foreach (BatchConfigSection.ProcessParamsElementCollection.ParamElement.ReplaceAttributesElementCollection.ReplaceElement pe in p.ReplaceAttributes)
                            {
                                string ReplaceValue = pe.ReplaceValue;
                                
                                if (n.Attributes[pe.AttributeNameForDashReplace] != null)
                                    ReplaceValue = ReplaceValue.Replace("#", n.Attributes[pe.AttributeNameForDashReplace].Value);
                                else ReplaceValue = ReplaceValue.Replace("#", pe.DefaultValueForDashReplace);

                                if (pe.AttributeName == "")
                                {
                                    string t;
                                    t = Regex.Replace(n.InnerText, pe.OriginalValue, ReplaceValue);
                                    if (t != n.InnerText)
                                    {
                                        n.InnerText = t;
                                        ChangesCount += 1;
                                    }
                                }
                                else 
                                foreach(XmlAttribute a in n.Attributes)
                                {
                                    if (Regex.IsMatch(a.Name, pe.AttributeName, RegexOptions.IgnoreCase))
                                    {
                                        string t;
                                        t = Regex.Replace(a.Value, pe.OriginalValue, ReplaceValue);
                                        if (t != a.Value)
                                        {
                                            a.Value = t;
                                            ChangesCount += 1;
                                        }
                                    }                                    
                                }                                
                            }                           
                         


                            foreach (BatchConfigSection.ProcessParamsElementCollection.ParamElement.DeduplicateElementsElementCollection.DeduplicateElement pe in p.DeduplicateElements)
                            {
                                Dictionary<string, int> dictionary = new Dictionary<string, int>();                                

                                for (int i = 0; i < n.ChildNodes.Count; i++)
                                {                                    
                                    if(n.ChildNodes.Item(i).Name == pe.XMLElement)
                                    {
                                        if (dictionary.ContainsKey(n.ChildNodes.Item(i).Attributes[pe.XMLKey].Value))
                                        {
                                            ChangesCount += 1;

                                            if (pe.Mode == "Rename")
                                            {
                                                string s = n.ChildNodes.Item(i).Attributes[pe.XMLKey].Value;

                                                if (s[s.Length-1] == ']') s = s.Insert(s.Length - 1, i.ToString());
                                                else s += i.ToString();

                                                n.ChildNodes.Item(i).Attributes[pe.XMLKey].Value = s;
                                            }
                                            else
                                                if (pe.Mode == "Remove")
                                                {
                                                    n.ChildNodes.Item(i).Attributes[pe.XMLKey].Value += i.ToString();

                                                    //((XmlElement)(n.ChildNodes.Item(i))).SetAttribute("DEBUG", "DELETE0-"+i.ToString());
                                                    n.RemoveChild(n.ChildNodes.Item(i));

                                                    for (int c = 0; c < pe.CascadeRemoves.Count; c++)
                                                    {
                                                        if (pe.CascadeRemoves[c].CascadeFrom == -1) ChangesCount += CascadeRemove(XMLDoc.DocumentElement, nsmgr, i, pe.CascadeRemoves, c);
                                                    }

                                                }
                                        }
                                        else
                                            dictionary.Add(n.ChildNodes.Item(i).Attributes[pe.XMLKey].Value, 0);
                                    }
                                }                                
                            }

                            foreach (BatchConfigSection.ProcessParamsElementCollection.ParamElement.ExtractContentsElementCollection.Extract pe in p.ExtractContents)
                            {
                                if (ExtractWriter == null)
                                {
                                    string ExtractsDirectory = Path.GetDirectoryName(ExtractsFileName);

                                    if (!Directory.Exists(ExtractsDirectory))
                                    {
                                        Directory.CreateDirectory(ExtractsDirectory);
                                    }

                                    ExtractWriter = new StreamWriter(ExtractsFileName);
                                    ExtractWriter.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                                    ExtractWriter.Write("<body xmlns:" + p.NodeXPath[1] + "=\"" + p.NodeNamespace + "\">");
                                }
                                ExtractContents(ExtractWriter, n, nsmgr, pe.SliceByXPaths);
                            }
                        }     
                    }
            }

            if (ChangesCount > 0) Console.WriteLine(ChangesCount.ToString() + " change(s) made.");

            if (XMLDoc == null) e.Extract(dest); else XMLDoc.Save(dest);
            dest.Position = 0;

            if (ExtractWriter != null)
            {
                ExtractWriter.Write("</body>");
                ExtractWriter.Flush();
                ExtractWriter.Close();
                ExtractWriter.Dispose();
            }

            return ChangesCount;
        }

        public void Dispose()
        {
        }
    }
}
