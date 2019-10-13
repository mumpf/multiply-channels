using System;
using System.Xml;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace MultiplyChannels {
    public class ProcessInclude {

        const string cOwnNamespace = "http://github.com/mumpf/multiply-channels";
        private XmlNamespaceManager nsmgr;
        private XmlDocument mDocument = new XmlDocument();
        private bool mLoaded = false;
        StringBuilder mHeaderGenerated = new StringBuilder();

        private XmlNode mParameterTypesNode = null;
        private static Dictionary<string, ProcessInclude> gIncludes = new Dictionary<string, ProcessInclude>();
        private string mXmlFileName;
        private string mHeaderFileName;
        private string mHeaderPrefixName;
        private bool mHeaderParameterStartGenerated;
        private bool mHeaderParameterBlockGenerated;
        private bool mHeaderKoStartGenerated;
        private bool mHeaderKoBlockGenerated;
        private int mChannelCount = 1;
        private int mParameterBlockOffset = 0;
        private int mParameterBlockSize = -1;
        private int mKoOffset = 0;
        private int mKoBlockSize = 0;

        public int ParameterBlockOffset {
            get { return mParameterBlockOffset; }
            set { mParameterBlockOffset = value; }
        }

        public int ParameterBlockSize {
            get { return mParameterBlockSize; }
            set { mParameterBlockSize = value; }
        }
        public int ChannelCount {
            get { return mChannelCount; }
            set { mChannelCount = value; }
        }

        public int KoOffset {
            get { return mKoOffset; }
            set { mKoOffset = value; }
        }

        public string HeaderGenerated {
            get {
                mHeaderGenerated.Insert(0, "#pragma once\n#include <knx.h>\n\n");
                return mHeaderGenerated.ToString();
            }
        }

        public static ProcessInclude Factory(string iXmlFileName, string iHeaderFileName, string iHeaderPrefixName) {
            ProcessInclude lInclude = null;
            if (gIncludes.ContainsKey(iXmlFileName)) {
                lInclude = gIncludes[iXmlFileName];
            } else {
                lInclude = new ProcessInclude(iXmlFileName, iHeaderFileName, iHeaderPrefixName);
                gIncludes.Add(iXmlFileName, lInclude);
            }
            return lInclude;
        }

        private ProcessInclude(string iXmlFileName, string iHeaderFileName, string iHeaderPrefixName) {
            mXmlFileName = iXmlFileName;
            mHeaderFileName = iHeaderFileName;
            mHeaderPrefixName = iHeaderPrefixName;
        }

        int GetHeaderParameter(string iHeaderFileContent, string iDefineName) {
            string lPattern = "#define.*" + iDefineName + @"\s*(\d{1,4})";
            Match m = Regex.Match(iHeaderFileContent, lPattern, RegexOptions.None);
            int lResult = -1;
            if (m.Groups.Count > 1) {
                int.TryParse(m.Groups[1].Value, out lResult);
            }
            return lResult;
        }

        bool ParseHeaderFile(string iHeaderFileName) {
            if (File.Exists(iHeaderFileName)) {
                StreamReader lHeaderFile = File.OpenText(iHeaderFileName);
                string lHeaderFileContent = lHeaderFile.ReadToEnd();
                lHeaderFile.Close();
                mChannelCount = GetHeaderParameter(lHeaderFileContent, mHeaderPrefixName + "_Channels");
                mKoOffset = GetHeaderParameter(lHeaderFileContent, mHeaderPrefixName + "_KoOffset");
            } else {
                mChannelCount = 1;
                mKoOffset = 1;
            }
            // mKoBlockSize = GetHeaderParameter(lHeaderFileContent, mHeaderPrefixName + "_KoBlockSize");
            return (mChannelCount >= 0) && (mKoOffset > 0);
        }

        public XmlNodeList SelectNodes(string iXPath) {
            return mDocument.SelectNodes(iXPath);
        }


        static string CalculateId(int iApplicationNumber, int iApplicationVersion) {
            return string.Format("-{0:X4}-{1:X2}-0000", iApplicationNumber, iApplicationVersion);
        }

        public XmlDocument GetDocument() {
            return mDocument;
        }

        public void DocumentDebugOutput() {
            mDocument.Save(Path.ChangeExtension(mXmlFileName, "out.xml"));
        }

        public void SetNamespace() {
            // we restor the original namespace, if necessary
            if (mDocument.DocumentElement.GetAttribute("xmlns") == "") {
                string lXmlns = mDocument.DocumentElement.GetAttribute("oldxmlns");
                if (lXmlns != "") mDocument.DocumentElement.SetAttribute("xmlns", lXmlns);
            }
        }

        public void Expand() {
            // here we recursively process all includes and all channel repetitions
            LoadAdvanced(mXmlFileName);
            ExportHeader(mHeaderFileName, mHeaderPrefixName, this);
            // finally we do all processing necessary for the whole (resolved) document
            ProcessFinish(mDocument);
            // DocumentDebugOutput();
        }

        string ReplaceKoTemplate(string iValue, int iChannel, ProcessInclude iInclude) {
            string lResult = iValue;
            Match lMatch = Regex.Match(iValue, @"%K(\d{1,3})%");
            int lBlockSize = 0;
            int lOffset = 0;
            if (iInclude != null) {
                lBlockSize = iInclude.mKoBlockSize;
                lOffset = iInclude.KoOffset;
            }
            // MatchCollection lMatches = Regex.Matches(iValue, @"%K(\d{1,3})%");
            if (lMatch.Captures.Count > 0) {
                int lShift = int.Parse(lMatch.Groups[1].Value);
                lResult = iValue.Replace(lMatch.Value, ((iChannel - 1) * lBlockSize + lOffset + lShift).ToString());
            }
            return lResult;
        }

        void ProcessAttributes(int iChannel, XmlNode iTargetNode, ProcessInclude iInclude) {
            foreach (XmlAttribute lAttr in iTargetNode.Attributes) {
                lAttr.Value = lAttr.Value.Replace("%C%", iChannel.ToString());
                lAttr.Value = ReplaceKoTemplate(lAttr.Value, iChannel, iInclude);
                // lAttr.Value = lAttr.Value.Replace("%N%", mChannelCount.ToString());
            }
        }

        void ProcessParameter(int iChannel, XmlNode iTargetNode, ProcessInclude iInclude) {
            //calculate new offset
            XmlNode lMemory = iTargetNode.SelectSingleNode("Memory");
            if (lMemory != null) {
                XmlNode lAttr = lMemory.Attributes.GetNamedItem("Offset");
                int lOffset = int.Parse(lAttr.Value);
                lOffset += iInclude.ParameterBlockOffset + (iChannel - 1) * iInclude.ParameterBlockSize;
                lAttr.Value = lOffset.ToString();
            }
        }

        void ProcessChannel(int iChannel, XmlNode iTargetNode, ProcessInclude iInclude) {
            //attributes of the node
            if (iTargetNode.Attributes != null) {
                ProcessAttributes(iChannel, iTargetNode, iInclude);
            }

            //Print individual children of the node, gets only direct children of the node
            XmlNodeList lChildren = iTargetNode.ChildNodes;
            foreach (XmlNode lChild in lChildren) {
                ProcessChannel(iChannel, lChild, iInclude);
            }
        }


        void ProcessTemplate(int iChannel, XmlNode iTargetNode, ProcessInclude iInclude) {
            ProcessAttributes(iChannel, iTargetNode, iInclude);
            if (iTargetNode.Name == "Parameter") {
                ProcessParameter(iChannel, iTargetNode, iInclude);
            } else
            if (iTargetNode.Name == "Channel") {
                ProcessChannel(iChannel, iTargetNode, iInclude);
            }
        }

        void ProcessIncludeFinish(XmlNode iTargetNode) {
            // set number of Channels
            XmlNodeList lNodes = iTargetNode.SelectNodes("//*[@Value='%N%']");
            foreach (XmlNode lNode in lNodes) {
                lNode.Attributes.GetNamedItem("Value").Value = mChannelCount.ToString();
            }
            lNodes = iTargetNode.SelectNodes("//*[@maxInclusive='%N%']");
            foreach (XmlNode lNode in lNodes) {
                lNode.Attributes.GetNamedItem("maxInclusive").Value = mChannelCount.ToString();
            }
            // // set the max channel value
            // ReplaceDocumentStrings(mDocument, "%N%", mChannelCount.ToString());
        }

        void ReplaceDocumentStrings(XmlNodeList iNodeList, string iSourceText, string iTargetText) {
            foreach (XmlNode lNode in iNodeList) {
                if (lNode.Attributes != null) {
                    foreach (XmlNode lAttribute in lNode.Attributes) {
                        lAttribute.Value = lAttribute.Value.ToString().Replace(iSourceText, iTargetText);
                    }
                }
                ReplaceDocumentStrings(lNode, iSourceText, iTargetText);
            }
        }
        void ReplaceDocumentStrings(XmlNode iNode, string iSourceText, string iTargetText) {
            ReplaceDocumentStrings(iNode.ChildNodes, iSourceText, iTargetText);
        }

        void ReplaceDocumentStrings(string iSourceText, string iTargetText) {
            ReplaceDocumentStrings(mDocument.ChildNodes, iSourceText, iTargetText);
        }

        void ProcessFinish(XmlNode iTargetNode) {
            // set the right Size attributes
            XmlNodeList lNodes = iTargetNode.SelectNodes("(//RelativeSegment | //LdCtrlRelSegment | //LdCtrlWriteRelMem)[@Size]");
            // string lSize = (mChannelCount * mParameterBlockSize + mParameterBlockOffset).ToString();
            string lSize = mParameterBlockSize.ToString();
            foreach (XmlNode lNode in lNodes) {
                lNode.Attributes.GetNamedItem("Size").Value = lSize;
            }
            Console.WriteLine("- Final parameter size is {0}", lSize);
            // change all Id-Attributes / renumber ParameterSeparator and ParameterBlock
            XmlNode lApplicationProgramNode = iTargetNode.SelectSingleNode("/KNX/ManufacturerData/Manufacturer/ApplicationPrograms/ApplicationProgram");
            string lApplicationId = lApplicationProgramNode.Attributes.GetNamedItem("Id").Value;
            int lApplicationNumber = int.Parse(lApplicationProgramNode.Attributes.GetNamedItem("ApplicationNumber").Value);
            int lApplicationVersion = int.Parse(lApplicationProgramNode.Attributes.GetNamedItem("ApplicationVersion").Value);
            XmlNode lReplacesVersionsAttribute = lApplicationProgramNode.Attributes.GetNamedItem("ReplacesVersions");
            string lOldId = lApplicationId.Replace("M-00FA_A", ""); // CalculateId(1, 1);
            string lNewId = CalculateId(lApplicationNumber, lApplicationVersion);
            int lParameterSeparatorCount = 1;
            int lParameterBlockCount = 1;
            XmlNodeList lAttrs = iTargetNode.SelectNodes("//*/@*[string-length() > '13']");
            foreach (XmlNode lAttr in lAttrs) {
                if (lAttr.Value != null) {
                    lAttr.Value = lAttr.Value.Replace(lOldId, lNewId);
                    // ParameterSeparator is renumbered
                    if (lAttr.Value.Contains("_PS-")) {
                        lAttr.Value = string.Format("{0}-{1}", lAttr.Value.Substring(0, lAttr.Value.LastIndexOf('-')), lParameterSeparatorCount);
                        lParameterSeparatorCount += 1;
                    }
                    // ParameterBlock is renumbered
                    if (lAttr.Value.Contains("_PB-")) {
                        lAttr.Value = string.Format("{0}-{1}", lAttr.Value.Substring(0, lAttr.Value.LastIndexOf('-')), lParameterBlockCount);
                        lParameterBlockCount += 1;
                    }
                }
            }
            Console.WriteLine("- ApplicationNumber: {0}, ApplicationVersion: {1}, old ID is: {3}, new (calculated) ID is: {2}", lApplicationNumber, lApplicationVersion, lNewId, lOldId);

            // create registration entry
            XmlNode lHardwareVersionAttribute = iTargetNode.SelectSingleNode("/KNX/ManufacturerData/Manufacturer/Hardware/Hardware/@VersionNumber");
            int lHardwareVersion = int.Parse(lHardwareVersionAttribute.Value);
            XmlNode lRegistrationNumber = iTargetNode.SelectSingleNode("/KNX/ManufacturerData/Manufacturer/Hardware/Hardware/Hardware2Programs/Hardware2Program/RegistrationInfo/@RegistrationNumber");
            lRegistrationNumber.Value = string.Format("0001/{0}{1}", lHardwareVersion, lApplicationVersion);
            Console.WriteLine("- RegistrationVersion is: {0}", lRegistrationNumber.Value);

            // Add ReplacesVersions 
            if (lReplacesVersionsAttribute != null) {
                string lReplacesVersions = lReplacesVersionsAttribute.Value;
                Console.WriteLine("- ReplacesVersions entry is: {0}", lReplacesVersions);
                // string lOldVerion = string.Format(" {0}", lApplicationVersion - 1);
                // if (!lReplacesVersions.Contains(lOldVerion) && lReplacesVersions != (lApplicationVersion - 1).ToString()) lReplacesVersionsAttribute.Value += lOldVerion;
            }

        }

        public int CalcParamSize(XmlNode iParameter, XmlNode iParameterTypesNode) {
            int lResult = 0;
            if (iParameterTypesNode != null) {
                XmlNode lMemory = iParameter.SelectSingleNode("Memory");
                if (lMemory != null) {
                    // we calcucalte the size only, if the parameter uses some memory in the device storage
                    string lParameterTypeId = iParameter.Attributes.GetNamedItem("ParameterType").Value;
                    XmlNode lParameterType = iParameterTypesNode.SelectSingleNode(string.Format("ParameterType[@Id='{0}']", lParameterTypeId));
                    if (lParameterType != null) {
                        XmlNode lSizeInBitAttribute = lParameterType.SelectSingleNode("*/@SizeInBit");
                        if (lSizeInBitAttribute != null) {
                            lResult = (int.Parse(lSizeInBitAttribute.Value) - 1) / 8 + 1;
                        } else if (lParameterType.SelectSingleNode("TypeFloat") != null) {
                            lResult = 4;
                        }
                    }
                }
            }
            return lResult;
        }

        public int CalcParamSize(XmlNodeList iParameterList, XmlNode iParameterTypesNode) {
            int lResult = 0;
            foreach (XmlNode lNode in iParameterList) {
                int lSize = CalcParamSize(lNode, iParameterTypesNode);
                if (lSize > 0) {
                    // at this point we know there is a memory reference, we look at the offset
                    XmlNode lOffset = lNode.SelectSingleNode("*/@Offset");
                    lResult = Math.Max(lResult, int.Parse(lOffset.Value) + lSize);
                }
            }
            return lResult;
        }

        private string ReplaceChannelName(string iName) {
            string lResult = iName;
            // if (iName.Contains("%C%")) lResult = iName.Remove(0, iName.IndexOf("%C%") + 3);
            lResult = iName.Replace("%C%", "");
            lResult = lResult.Replace(" ", "_");
            return lResult;
        }

        public void ExportHeaderKoStart(StringBuilder cOut, string iHeaderPrefixName) {
            if (!mHeaderKoStartGenerated) {
                StringBuilder lOut = new StringBuilder();
                mHeaderKoStartGenerated = ExportHeaderKo(lOut, iHeaderPrefixName);
                if (mHeaderKoStartGenerated) {
                    cOut.AppendLine("// Communication objects with single occurance");
                    cOut.Append(lOut);
                }
            }
        }

        public void ExportHeaderKoBlock(StringBuilder cOut, string iHeaderPrefixName) {
            if (!mHeaderKoBlockGenerated) {
                XmlNodeList lComObjects = mDocument.SelectNodes("//ComObjectTable/ComObject");
                mKoBlockSize = lComObjects.Count;

                StringBuilder lOut = new StringBuilder();
                mHeaderKoBlockGenerated = ExportHeaderKo(lOut, iHeaderPrefixName);
                if (mHeaderKoBlockGenerated) {
                    cOut.AppendLine("// Communication objects per channel (multiple occurance)");
                    cOut.AppendFormat("#define {0}_KoOffset {1}", iHeaderPrefixName, mKoOffset);
                    cOut.AppendLine();
                    cOut.AppendFormat("#define {0}_KoBlockSize {1}", iHeaderPrefixName, mKoBlockSize);
                    cOut.AppendLine();
                    cOut.Append(lOut);
                }
            }
        }

        public bool ExportHeaderKo(StringBuilder cOut, string iHeaderPrefixName) {
            XmlNodeList lNodes = mDocument.SelectNodes("//ComObject");
            bool lResult = false;
            foreach (XmlNode lNode in lNodes) {
                string lNumber = ReplaceKoTemplate(lNode.Attributes.GetNamedItemValueOrEmpty("Number"), 1, null);
                cOut.AppendFormat("#define {0}_Ko{1} {2}", iHeaderPrefixName, ReplaceChannelName(lNode.NodeName()), lNumber);
                cOut.AppendLine();
                lResult = true;
            }
            if (lResult) cOut.AppendLine();
            return lResult;
        }

        public void ExportHeaderParameterStart(StringBuilder cOut, XmlNode iParameterTypesNode, string iHeaderPrefixName) {
            if (!mHeaderParameterStartGenerated) {
                cOut.AppendLine("// Parameter with single occurance");
                ExportHeaderParameter(cOut, iParameterTypesNode, iHeaderPrefixName);
                mHeaderParameterStartGenerated = true;
            }
        }

        public void ExportHeaderParameterBlock(StringBuilder cOut, XmlNode iParameterTypesNode, string iHeaderPrefixName) {
            if (!mHeaderParameterBlockGenerated) {
                cOut.AppendFormat("#define {0}_Channels {1}", iHeaderPrefixName, mChannelCount);
                cOut.AppendLine();
                cOut.AppendLine();
                cOut.AppendLine("// Parameter per channel");
                cOut.AppendFormat("#define {0}_ParamBlockOffset {1}", iHeaderPrefixName, mParameterBlockOffset);
                cOut.AppendLine();
                cOut.AppendFormat("#define {0}_ParamBlockSize {1}", iHeaderPrefixName, mParameterBlockSize);
                cOut.AppendLine();
                int lSize = ExportHeaderParameter(cOut, iParameterTypesNode, iHeaderPrefixName);
                // if (lSize != mParameterBlockSize) throw new ArgumentException(string.Format("ParameterBlockSize {0} calculation differs from header filie calculated ParameterBlockSize {1}", mParameterBlockSize, lSize));
                mHeaderParameterBlockGenerated = true;
            }
        }

        public int ExportHeaderParameter(StringBuilder cOut, XmlNode iParameterTypesNode, string iHeaderPrefixName) {
            int lMaxSize = 0;
            XmlNodeList lNodes = mDocument.SelectNodes("//Parameter");
            foreach (XmlNode lNode in lNodes) {
                string lName = lNode.Attributes.GetNamedItem("Name").Value;
                lName = ReplaceChannelName(lName);
                XmlNode lMemory = lNode.FirstChild;
                while (lMemory != null && lMemory.NodeType == XmlNodeType.Comment) lMemory = lMemory.NextSibling;
                if (lMemory != null && iParameterTypesNode != null) {
                    // parse parameter type to fill additional information
                    string lParameterTypeId = lNode.Attributes.GetNamedItem("ParameterType").Value;
                    XmlNode lParameterType = iParameterTypesNode.SelectSingleNode(string.Format("//ParameterType[@Id='{0}']", lParameterTypeId));
                    XmlNode lTypeNumber = null;
                    if (lParameterType != null) lTypeNumber = lParameterType.FirstChild;
                    while (lTypeNumber != null && lTypeNumber.NodeType == XmlNodeType.Comment) lTypeNumber = lTypeNumber.NextSibling;
                    int lBits = 0;
                    string lType = "";
                    bool lDirectType = false;
                    if (lTypeNumber != null) {
                        XmlNode lBitsAttribute = lTypeNumber.Attributes.GetNamedItem("SizeInBit");
                        if (lBitsAttribute != null) lBits = int.Parse(lBitsAttribute.Value);
                        XmlNode lTypeAttribute = lTypeNumber.Attributes.GetNamedItem("Type");
                        if (lTypeAttribute != null) {
                            lType = lTypeAttribute.Value;
                            lType = (lType == "signedInt") ? "int" : (lType == "unsignedInt") ? "uint" : "xxx";
                        } else {
                            lType = "enum";
                            if (lBits > 8) {
                                lType = string.Format("char*, {0} Byte", lBits / 8);
                                lDirectType = true;
                            }
                        }
                        if (lTypeNumber.Name == "TypeFloat") {
                            lType = "float";
                            lBits = 16;
                            lDirectType = true;
                        }
                    }
                    int lOffset = int.Parse(lMemory.Attributes.GetNamedItem("Offset").Value);
                    int lBitOffset = int.Parse(lMemory.Attributes.GetNamedItem("BitOffset").Value);
                    lMaxSize = Math.Max(lMaxSize, lOffset + (lBits - 1) / 8 + 1);
                    if (lBits <= 7 || lType == "enum") {
                        //output for bit based parameters 
                        lType = string.Format("{0} Bit{1}, Bit {2}", lBits, (lBits == 1) ? "" : "s", (7 - lBitOffset));
                        if (lBits > 1) lType = string.Format("{0}-{1}", lType, 8 - lBits - lBitOffset);
                        cOut.AppendFormat("#define {3}_{0,-25} {1,2}      // {2}", lName, lOffset, lType, iHeaderPrefixName);
                    } else if (lDirectType) {
                        cOut.AppendFormat("#define {3}_{0,-25} {1,2}      // {2}", lName, lOffset, lType, iHeaderPrefixName);
                    } else {
                        cOut.AppendFormat("#define {4}_{0,-25} {1,2}      // {3}{2}_t", lName, lOffset, lBits, lType, iHeaderPrefixName);
                    }
                    cOut.AppendLine();
                }
            }
            cOut.AppendLine();
            return lMaxSize;
        }


        /// <summary>
        /// Load xml document from file resolving xincludes recursivly
        /// </summary>
        public void LoadAdvanced(string iFileName) {
            if (!mLoaded) {
                string lCurrentDir = Path.GetDirectoryName(Path.GetFullPath(iFileName));
                string lFileData = File.ReadAllText(iFileName);
                if (lFileData.Contains("oldxmlns")) {
                    // we get rid of default namespace, we already have an original (this file was already processed by our processor)
                    int lStart = lFileData.IndexOf(" xmlns=\"");
                    if (lStart < 0) {
                        lFileData = lFileData.Replace("oldxmlns", "xmlns");
                    } else {
                        // int lEnd = lFileData.IndexOf("\"", lStart + 8) + 1;
                        lFileData = lFileData.Remove(lStart, 38);
                        // lFileData = lFileData.Substring(0, lStart) + lFileData.Substring(lEnd);
                    }
                } else {
                    // we get rid of default namespace, but remember the original
                    lFileData = lFileData.Replace(" xmlns=\"", " oldxmlns=\"");
                }
                using (StringReader sr = new StringReader(lFileData)) {
                    mDocument.Load(sr);
                    mLoaded = true;
                    ResolveIncludes(lCurrentDir);
                }
            }
        }

        /// <summary>
        /// Resolves Includes inside xml document
        /// </summary>
        /// <param name="iCurrentDir">Directory to use for relative href expressions</param>
        public void ResolveIncludes(string iCurrentDir) {
            nsmgr = new XmlNamespaceManager(mDocument.NameTable);
            nsmgr.AddNamespace("mc", cOwnNamespace);

            //find all XIncludes in a copy of the document
            XmlNodeList lIncludeNodes = mDocument.SelectNodes("//mc:include", nsmgr); // get all <include> nodes

            foreach (XmlNode lIncludeNode in lIncludeNodes)
            // try
            {
                //Load document...
                string lIncludeName = lIncludeNode.Attributes.GetNamedItemValueOrEmpty("href");
                string lHeaderFileName = lIncludeNode.Attributes.GetNamedItemValueOrEmpty("header");
                string lHeaderPrefixName = lIncludeNode.Attributes.GetNamedItemValueOrEmpty("prefix");
                ProcessInclude lInclude = ProcessInclude.Factory(lIncludeName, lHeaderFileName, lHeaderPrefixName);
                string lTargetPath = Path.Combine(iCurrentDir, lIncludeName);
                lInclude.LoadAdvanced(lTargetPath);
                //...find include in real document...
                XmlNode lParent = lIncludeNode.ParentNode;
                string lXPath = lIncludeNode.Attributes.GetNamedItemValueOrEmpty("xpath");
                XmlNodeList lChildren = lInclude.SelectNodes(lXPath);
                if (lHeaderFileName != "") {
                    lHeaderFileName = Path.Combine(iCurrentDir, lHeaderFileName);
                    if (lChildren.Count > 0 && "Parameter | ComObject".Contains(lChildren[0].LocalName)) {
                        // at this point we are including a template file
                        ExportHeader(lHeaderFileName, lHeaderPrefixName, lInclude, lChildren);
                    }
                }
                // here we do template processing and repeat the template as many times as
                // the Channels parameter in header file
                for (int lChannel = 1; lChannel <= lInclude.ChannelCount; lChannel++) {
                    foreach (XmlNode lChild in lChildren) {
                        //necessary for move between XmlDocument contexts
                        XmlNode lImportNode = lParent.OwnerDocument.ImportNode(lChild, true);
                        // for any Parameter node we do offset recalculation
                        // if there is no prefixname, we do no template replacement
                        if (lHeaderPrefixName != "") ProcessTemplate(lChannel, lImportNode, lInclude);
                        lParent.InsertBefore(lImportNode, lIncludeNode);
                    }
                }
                lParent.RemoveChild(lIncludeNode);
                if (lInclude.ChannelCount > 1) ReplaceDocumentStrings("%N%", lInclude.ChannelCount.ToString());
                // if (lHeaderPrefixName != "") ProcessIncludeFinish(lChildren);
                //if this fails, something is wrong
            }
            // catch { }
        }

        public void ExportHeader(string iHeaderFileName, string iHeaderPrefixName, ProcessInclude iInclude, XmlNodeList iChildren = null) {
            iInclude.ParseHeaderFile(iHeaderFileName);

            if (mParameterTypesNode == null) {
                // before we start with template processing, we calculate all Parameter relevant info
                mParameterTypesNode = mDocument.SelectSingleNode("//ParameterTypes");
            }

            if (mParameterTypesNode != null) {
                // the main document contains necessary ParameterTypes definitions
                // there are new parameters in include, we have to calculate a new parameter offset
                XmlNodeList lParameterNodes = mDocument.SelectNodes("//Parameters/Parameter");
                if (lParameterNodes != null) {
                    mParameterBlockSize = CalcParamSize(lParameterNodes, mParameterTypesNode);
                }
                if (iChildren != null) {
                    // ... and we do parameter processing, so we calculate ParamBlockSize for this include
                    int lBlockSize = iInclude.CalcParamSize(iChildren, mParameterTypesNode);
                    if (lBlockSize > 0) {
                        iInclude.ParameterBlockSize = lBlockSize;
                        // we calculate also ParamOffset
                        iInclude.ParameterBlockOffset = mParameterBlockSize;
                    }
                }
            }
            // Header file generation is only possible before we resolve includes
            // First we serialize local parameters of this instance
            ExportHeaderParameterStart(mHeaderGenerated, mParameterTypesNode, iHeaderPrefixName);
            // followed by template parameters of the include
            if (iInclude != this) iInclude.ExportHeaderParameterBlock(mHeaderGenerated, mParameterTypesNode, iHeaderPrefixName);

            ExportHeaderKoStart(mHeaderGenerated, iHeaderPrefixName);
            if (iInclude != this) iInclude.ExportHeaderKoBlock(mHeaderGenerated, iHeaderPrefixName);
        }
    }
}

