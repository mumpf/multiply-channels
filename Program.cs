using System;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using CommandLine;
using System.Linq;
// using Knx.Ets.Xml.ObjectModel;

//using Knx.Ets.Converter.ConverterEngine;

namespace MultiplyChannels {

    struct EtsVersion {
        public EtsVersion(string iSubdir, string iEts) {
            Subdir = iSubdir;
            ETS = iEts;
        }

        public string Subdir { get; private set; }
        public string ETS { get; private set; }
    }
    class Program {

        private static Dictionary<string, EtsVersion> EtsVersions = new Dictionary<string, EtsVersion>() {
            {"http://knx.org/xml/project/11", new EtsVersion(@"CV\4.0.1997.50261", "ETS 4")},
            {"http://knx.org/xml/project/12", new EtsVersion(@"CV\5.0.204.12971", "ETS 5")},
            {"http://knx.org/xml/project/13", new EtsVersion(@"CV\5.1.84.17602", "ETS 5.5")},
            {"http://knx.org/xml/project/14", new EtsVersion(@"CV\5.6.241.33672", "ETS 5.6")}
        };

        private const string gtoolName = "KNX MT";
        private const string gtoolVersion = "5.1.255.16695";
        //installation path of a valid ETS instance (only ETS4 or ETS5 supported)
        private static string gPathETS = @"C:\Program Files (x86)\ETS5";

        static string FindEtsPath(string iXmlFilename) {
            string lResult = "";
            if (File.Exists(iXmlFilename)) {
                string lXmlContent = File.ReadLines(iXmlFilename).Take(2).Aggregate((s1, s2) => s1 + s2);
                int lStart = lXmlContent.IndexOf("xmlns=\"http://knx.org/xml/project/", 0);
                string lXmlns = lXmlContent.Substring(lStart + 7, 29);
                int lProjectVersion = int.Parse(lXmlns.Substring(27));

                if (EtsVersions.ContainsKey(lXmlns)) {
                    string lSubdir = EtsVersions[lXmlns].Subdir;
                    string lEts = EtsVersions[lXmlns].ETS;
                    string lPath = Path.Combine(gPathETS, lSubdir);
                    if (!Directory.Exists(lPath)) {
                        // fallback handling
                        if (lProjectVersion == 11) {
                            // for project/11 ETS4 might be installed
                            lPath = gPathETS.Replace("5", "4");
                            lEts = "ETS 4";
                        } else {
                            // for other versions the required DLL might be in ETS installation dir
                            // in this case the CV dir contains the previous project conversion dlls
                            string lTempXmlns = string.Format("http://knx.org/xml/project/{0}", lProjectVersion - 1);
                            lSubdir = EtsVersions[lTempXmlns].Subdir;
                            lPath = Path.Combine(gPathETS, lSubdir);
                            if (Directory.Exists(lPath)) lPath = gPathETS;
                        }
                    }
                    if (Directory.Exists(lPath)) {
                        lResult = lPath;
                        Console.WriteLine("Found namespace {1} in xml, will use {0} for conversion...", lEts, lXmlns);
                    }
                }
                if (lResult == "") Console.WriteLine("No valid conversion engine available for xmlns {0}", lXmlns);
            }
            return lResult;
        }

        static void WriteFail(ref bool iFail, string iFormat, params object[] iParams) {
            if (!iFail) Console.WriteLine();
            Console.WriteLine("  --> " + iFormat, iParams);
            iFail = true;
        }

        // Parameter cache
        static Dictionary<string, XmlNode> gParameters = new Dictionary<string, XmlNode>();

        static XmlNode GetParameterById(XmlNode iRootNode, string iId) {
            XmlNode lResult = null;
            if (gParameters.ContainsKey(iId)) {
                lResult = gParameters[iId];
            } else {
                lResult = iRootNode.SelectSingleNode(string.Format("//Parameter[@Id='{0}']", iId));
                if (lResult != null) gParameters.Add(iId, lResult);
            }
            return lResult;
        }

        static bool ProcessSanityChecks(XmlNode iTargetNode) {

            Console.WriteLine();
            Console.WriteLine("Sanity checks... ");
            bool lFail = false;

            Console.Write("- Id-Uniqueness...");
            bool lFailPart = false;
            XmlNodeList lNodes = iTargetNode.SelectNodes("//*[@Id]");
            Dictionary<string, bool> lIds = new Dictionary<string, bool>();
            foreach (XmlNode lNode in lNodes) {
                string lId = lNode.Attributes.GetNamedItem("Id").Value;
                if (lIds.ContainsKey(lId)) {
                    WriteFail(ref lFailPart, "{0} is a duplicate Id in {1}", lId, lNode.NodeAttr("Name"));
                } else {
                    lIds.Add(lId, true);
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            lFailPart = false;
            Console.Write("- RefId-Integrity...");
            lNodes = iTargetNode.SelectNodes("//*[@RefId]");
            foreach (XmlNode lNode in lNodes) {
                if (lNode.Name != "Manufacturer") {
                    string lRefId = lNode.Attributes.GetNamedItem("RefId").Value;
                    if (!lIds.ContainsKey(lRefId)) {
                        WriteFail(ref lFailPart, "{0} is referenced in {1} {2}, but not defined", lRefId, lNode.Name, lNode.NodeAttr("Name"));
                    }
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            lFailPart = false;
            Console.Write("- ParamRefId-Integrity...");
            lNodes = iTargetNode.SelectNodes("//*[@ParamRefId]");
            foreach (XmlNode lNode in lNodes) {
                if (lNode.Name != "Manufacturer") {
                    string lParamRefId = lNode.Attributes.GetNamedItem("ParamRefId").Value;
                    if (!lIds.ContainsKey(lParamRefId)) {
                        WriteFail(ref lFailPart, "{0} is referenced in {1} {2}, but not defined", lParamRefId, lNode.Name, lNode.NodeAttr("Name"));
                    }
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            lFailPart = false;
            Console.Write("- TextParameterRefId-Integrity...");
            lNodes = iTargetNode.SelectNodes("//*[@TextParameterRefId]");
            foreach (XmlNode lNode in lNodes) {
                string lTextParamRefId = lNode.Attributes.GetNamedItem("TextParameterRefId").Value;
                if (!lIds.ContainsKey(lTextParamRefId)) {
                    WriteFail(ref lFailPart, "{0} is referenced in {1} {2}, but not defined", lTextParamRefId, lNode.Name, lNode.NodeAttr("Name"));
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            lFailPart = false;
            Console.Write("- ParameterType-Integrity...");
            lNodes = iTargetNode.SelectNodes("//*[@ParameterType]");
            foreach (XmlNode lNode in lNodes) {
                string lParameterType = lNode.Attributes.GetNamedItem("ParameterType").Value;
                if (!lIds.ContainsKey(lParameterType)) {
                    WriteFail(ref lFailPart, "{0} is referenced in {1} {2}, but not defined", lParameterType, lNode.Name, lNode.NodeAttr("Name"));
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            Console.Write("- Parameter-Name-Uniqueness...");
            lFailPart = false;
            lNodes = iTargetNode.SelectNodes("//Parameter[@Name]");
            Dictionary<string, bool> lParameterNames = new Dictionary<string, bool>();
            foreach (XmlNode lNode in lNodes) {
                string lName = lNode.Attributes.GetNamedItem("Name").Value;
                if (lParameterNames.ContainsKey(lName)) {
                    WriteFail(ref lFailPart, "{0} is a duplicate Name in Parameter '{1}'", lName, lNode.NodeAttr("Text"));
                } else {
                    lParameterNames.Add(lName, true);
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            lFailPart = false;
            Console.Write("- Parameter-Value-Integrity...");
            lNodes = iTargetNode.SelectNodes("//Parameter");
            foreach (XmlNode lNode in lNodes) {
                // we add the node to parameter cache
                string lNodeId = lNode.NodeAttr("Id");
                if (!gParameters.ContainsKey(lNodeId)) gParameters.Add(lNodeId, lNode);
                string lMessage = string.Format("Parameter {0}", lNode.NodeAttr("Name"));
                string lParameterValue = lNode.NodeAttr("Value", null);
                if (lParameterValue == null) {
                    WriteFail(ref lFailPart, "{0} has no Value attribute", lMessage);
                }
                lFailPart = CheckParameterValueIntegrity(iTargetNode, lFailPart, lNode, lParameterValue, lMessage);
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            lFailPart = false;
            Console.Write("- ParameterRef-Value-Integrity...");
            lNodes = iTargetNode.SelectNodes("//ParameterRef[@Value]");
            foreach (XmlNode lNode in lNodes) {
                string lParameterRefValue = lNode.NodeAttr("Value");
                // find parameter
                XmlNode lParameterNode = GetParameterById(iTargetNode, lNode.NodeAttr("RefId"));
                string lMessage = string.Format("ParameterRef {0}, referencing Parameter {1},", lNode.NodeAttr("Id"), lParameterNode.NodeAttr("Name"));
                lFailPart = CheckParameterValueIntegrity(iTargetNode, lFailPart, lParameterNode, lParameterRefValue, lMessage);
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            Console.Write("- ComObject-Name-Uniqueness...");
            lFailPart = false;
            lNodes = iTargetNode.SelectNodes("//ComObject[@Name]");
            Dictionary<string, bool> lKoNames = new Dictionary<string, bool>();
            foreach (XmlNode lNode in lNodes) {
                string lName = lNode.Attributes.GetNamedItem("Name").Value;
                if (lKoNames.ContainsKey(lName)) {
                    WriteFail(ref lFailPart, "{0} is a duplicate Name in ComObject number {1}", lName, lNode.NodeAttr("Number"));
                } else {
                    lKoNames.Add(lName, true);
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            Console.Write("- ComObject-Number-Uniqueness...");
            lFailPart = false;
            lNodes = iTargetNode.SelectNodes("//ComObject[@Number]");
            Dictionary<int, bool> lKoNumbers = new Dictionary<int, bool>();
            foreach (XmlNode lNode in lNodes) {
                int lNumber = int.Parse(lNode.Attributes.GetNamedItem("Number").Value);
                if (lKoNumbers.ContainsKey(lNumber)) {
                    WriteFail(ref lFailPart, "{0} is a duplicate Number in ComObject with name {1}", lNumber, lNode.NodeAttr("Name"));
                } else {
                    lKoNumbers.Add(lNumber, true);
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            Console.Write("- Id-Namespace...");
            // find refid
            lFailPart = false;
            XmlNode lApplicationProgramNode = iTargetNode.SelectSingleNode("/KNX/ManufacturerData/Manufacturer/ApplicationPrograms/ApplicationProgram");
            string lApplicationId = lApplicationProgramNode.Attributes.GetNamedItem("Id").Value;
            string lRefNs = lApplicationId.Replace("M-00FA_A", "");
            // check all nodes according to refid
            lNodes = iTargetNode.SelectNodes("//*/@*[string-length() > '13']");
            foreach (XmlNode lNode in lNodes) {
                if (lNode.Value != null) {
                    var lMatch = Regex.Match(lNode.Value, "-[0-9A-F]{4}-[0-9A-F]{2}-[0-9A-F]{4}");
                    if (lMatch.Success) {
                        if (lMatch.Value != lRefNs) {
                            XmlElement lElement = ((XmlAttribute)lNode).OwnerElement;
                            WriteFail(ref lFailPart, "{0} of node {2} {3} is in a different namespace than application namespace {1}", lMatch.Value, lRefNs, lElement.Name, lElement.NodeAttr("Name"));
                        }
                    }
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            lFailPart = false;
            Console.Write("- Serial number...");
            lNodes = iTargetNode.SelectNodes("//*[@SerialNumber]");
            foreach (XmlNode lNode in lNodes) {
                string lSerialNumber = lNode.Attributes.GetNamedItem("SerialNumber").Value;
                if (lSerialNumber.Contains("-")) {
                    WriteFail(ref lFailPart, "Hardware.SerialNumber={0}, it contains a dash (-), this will cause problems in knxprod.", lSerialNumber);
                }
            }
            if (!lFailPart) Console.WriteLine(" OK");
            lFail = lFail || lFailPart;

            return !lFail;
        }

        private static bool CheckParameterValueIntegrity(XmlNode iTargetNode, bool iFailPart, XmlNode iParameterNode, string iValue, string iMessage) {
            string lParameterType = iParameterNode.NodeAttr("ParameterType");
            if (lParameterType == "") {
                WriteFail(ref iFailPart, "Parameter {0} has no ParameterType attribute", iParameterNode.NodeAttr("Name"));
            }
            if (iValue != null && lParameterType != "") {
                // find parameter type
                XmlNode lParameterTypeNode = iTargetNode.SelectSingleNode(string.Format("//ParameterType[@Id='{0}']", lParameterType));
                if (lParameterTypeNode != null) {
                    // get first child ignoring comments
                    XmlNode lChild = lParameterTypeNode.ChildNodes[0];
                    while (lChild != null && lChild.NodeType != XmlNodeType.Element) lChild = lChild.NextSibling;
                    switch (lChild.Name) {
                        case "TypeNumber":
                            int lDummyInt;
                            bool lSuccess = int.TryParse(iValue, out lDummyInt);
                            if (!lSuccess) {
                                WriteFail(ref iFailPart, "Value of {0} cannot be converted to a number, value is '{1}'", iMessage, iValue);
                            }
                            break;
                        case "TypeFloat":
                            float lDummyFloat;
                            lSuccess = float.TryParse(iValue, out lDummyFloat);
                            if (!lSuccess || iValue.Contains(",")) {
                                WriteFail(ref iFailPart, "Value of {0} cannot be converted to a float, value is '{1}'", iMessage, iValue);
                            }
                            break;
                        case "TypeRestriction":
                            lSuccess = false;
                            foreach (XmlNode lEnumeration in lChild.ChildNodes) {
                                if (lEnumeration.Name == "Enumeration") {
                                    if (lEnumeration.NodeAttr("Value") == iValue) {
                                        lSuccess = true;
                                        break;
                                    }
                                }
                            }
                            if (!lSuccess) {
                                WriteFail(ref iFailPart, "Value of {0} is not contained in enumeration {2}, value is '{1}'", iMessage, iValue, lParameterType);
                            }
                            break;
                        default:
                            break;
                    }
                }
            }

            return iFailPart;
        }

        #region Reflection
        private static object InvokeMethod(Type type, string methodName, object[] args) {

            var mi = type.GetMethod(methodName, BindingFlags.Static | BindingFlags.NonPublic);
            return mi.Invoke(null, args);
        }

        private static void SetProperty(Type type, string propertyName, object value) {
            PropertyInfo prop = type.GetProperty(propertyName, BindingFlags.NonPublic | BindingFlags.Static);
            prop.SetValue(null, value, null);
        }
        #endregion

        // private static void ExportXsd(string iXsdFileName) {
        //     using (var fileStream = new FileStream("knx.xsd", FileMode.Create))
        //     using (var stream = DocumentSet.GetXmlSchemaDocumentAsStream(KnxXmlSchemaVersion.Version14)) {
        //         while (true) {
        //             var buffer = new byte[4096];
        //             var count = stream.Read(buffer, 0, 4096);
        //             if (count == 0)
        //                 break;

        //             fileStream.Write(buffer, 0, count);
        //         }
        //     }
        // }

        private static void ExportKnxprod(string iPathETS, string iXmlFileName, string iKnxprodFileName) {
            if (iPathETS == "") return;
            try {
                var files = new string[] { iXmlFileName };
                var asmPath = Path.Combine(iPathETS, "Knx.Ets.Converter.ConverterEngine.dll");
                var asm = Assembly.LoadFrom(asmPath);
                var eng = asm.GetType("Knx.Ets.Converter.ConverterEngine.ConverterEngine");
                var bas = asm.GetType("Knx.Ets.Converter.ConverterEngine.ConvertBase");

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                InvokeMethod(bas, "Uninitialize", null);
                //ConvertBase.Uninitialize();
                //var dset = ConverterEngine.BuildUpRawDocumentSet( files );
                var dset = InvokeMethod(eng, "BuildUpRawDocumentSet", new object[] { files });
                //ConverterEngine.CheckOutputFileName(outputFile, ".knxprod");
                InvokeMethod(eng, "CheckOutputFileName", new object[] { iKnxprodFileName, ".knxprod" });
                //ConvertBase.CleanUnregistered = false;
                //SetProperty(bas, "CleanUnregistered", false);
                //dset = ConverterEngine.ReOrganizeDocumentSet(dset);
                dset = InvokeMethod(eng, "ReOrganizeDocumentSet", new object[] { dset });
                //ConverterEngine.PersistDocumentSetAsXmlOutput(dset, outputFile, null, string.Empty, true, _toolName, _toolVersion);
                InvokeMethod(eng, "PersistDocumentSetAsXmlOutput", new object[] { dset, iKnxprodFileName, null, "", true, gtoolName, gtoolVersion });
                Console.WriteLine("Output of {0} successful", iKnxprodFileName);
            }
            catch (Exception ex) {
                Console.WriteLine("Error during knxprod creation:");
                Console.WriteLine(ex.ToString());
            }
        }

        class EtsOptions {
            private string mXmlFileName;
            [Value(0, MetaName = "xml file name", Required = true, HelpText = "Xml file name", MetaValue = "FILE")]
            public string XmlFileName {
                get { return mXmlFileName; }
                set { mXmlFileName = Path.ChangeExtension(value, "xml"); }
            }
        }

        [Verb("new", HelpText = "Create new xml file with a fully commented and working mini exaple")]
        class NewOptions : CreateOptions {
            [Option('x', "ProductName", Required = true, HelpText = "Product name - appears in catalog and in property dialog", MetaValue = "STRING")]
            public string ProductName { get; set; }
            [Option('n', "AppName", Required = false, HelpText = "(Default: Product name) Application name - appears in catalog and necessary for application upgrades", MetaValue = "STRING")]
            public string ApplicationName { get; set; } = "";
            [Option('a', "AppNumber", Required = true, HelpText = "Application number - has to be unique per manufacturer", MetaValue = "INT")]
            public int? ApplicationNumber { get; set; }
            [Option('y', "AppVersion", Required = false, Default = 1, HelpText = "Application version - necessary for application upgrades", MetaValue = "INT")]
            public int? ApplicationVersion { get; set; } = 1;
            [Option('w', "HardwareName", Required = false, HelpText = "(Default: Product name) Hardware name - not visible in ETS", MetaValue = "STRING")]
            public string HardwareName { get; set; } = "";
            [Option('v', "HardwareVersion", Required = false, Default = 1, HelpText = "Hardware version - not visible in ETS, required for registration", MetaValue = "INT")]
            public int? HardwareVersion { get; set; } = 1;
            [Option('s', "SerialNumber", Required = false, HelpText = "(Default: Application number) Hardware serial number - not visible in ETS, requered for hardware-id", MetaValue = "STRING")]
            public string SerialNumber { get; set; } = "";
            [Option('m', "MediumType", Required = false, Default = "TP", HelpText = "Medium type", MetaValue = "TP,IP,both")]
            public string MediumType { get; set; } = "TP";
            [Option('#', "OrderNumber", Required = false, HelpText = "(Default: Application number) Order number - appears in catalog and in property info tab", MetaValue = "STRING")]
            public string OrderNumber { get; set; } = "";

            public string MediumTypes {
                get {
                    string lResult = "MT-0";
                    if (MediumType == "IP") {
                        lResult = "MT-5";
                    } else if (MediumType == "both") {
                        lResult = "MT-0 MT-5";
                    }
                    return lResult;
                }
            }
            public string MaskVersion {
                get {
                    string lResult = "MV-07B0";
                    if (MediumType == "IP") lResult = "MV-57B0";
                    return lResult;
                }
            }
        }

        [Verb("knxprod", HelpText = "Create knxprod file from given xml file")]
        class KnxprodOptions : EtsOptions {
            [Option('o', "Output", Required = false, HelpText = "Output file name", MetaValue = "FILE")]
            public string OutputFile { get; set; } = "";
        }

        [Verb("create", HelpText = "Process given xml file with all includes and create knxprod")]
        class CreateOptions : KnxprodOptions {
            [Option('h', "HeaderFileName", Required = false, HelpText = "Header file name", MetaValue = "FILE")]
            public string HeaderFileName { get; set; } = "";
            [Option('p', "Prefix", Required = false, HelpText = "Prefix for generated contant names in header file", MetaValue = "STRING")]
            public string Prefix { get; set; } = "";
            [Option('d', "Debug", Required = false, HelpText = "Additional output of <xmlfile>.debug.xml, this file is the input file for knxprod converter")]
            public bool Debug { get; set; } = false;
        }

        [Verb("check", HelpText = "execute sanity checks on given xml file")]
        class CheckOptions : EtsOptions {
        }

        static int Main(string[] args) {
            return CommandLine.Parser.Default.ParseArguments<CreateOptions, CheckOptions, KnxprodOptions, NewOptions>(args)
              .MapResult(
                (NewOptions opts) => VerbNew(opts),
                (CreateOptions opts) => VerbCreate(opts),
                (KnxprodOptions opts) => VerbKnxprod(opts),
                (CheckOptions opts) => VerbCheck(opts),
                errs => 1);
        }

        static private int VerbNew(NewOptions opts) {
            // Handle defaults
            if (opts.ApplicationName == "") opts.ApplicationName = opts.ProductName;
            if (opts.HardwareName == "") opts.HardwareName = opts.ProductName;
            if (opts.SerialNumber == "") opts.SerialNumber = opts.ApplicationNumber.ToString();
            if (opts.OrderNumber == "") opts.OrderNumber = opts.ApplicationNumber.ToString();

            // checks
            bool lFail = false;
            if (opts.ApplicationNumber > 65535) {
                Console.WriteLine("ApplicationNumber has to be less than 65536!");
                lFail = true;
            }
            if (opts.SerialNumber.Contains("-")) {
                Console.WriteLine("SerialNumber must not contain a dash (-) character!");
                lFail = true;
            }
            if (opts.OrderNumber.Contains("-")) {
                Console.WriteLine("OrderNumber must not contain a dash (-) character!");
                lFail = true;
            }
            if (lFail) return 1;

            // create initial xml file
            string lXmlFile = "";
            var assembly = Assembly.GetEntryAssembly();
            var resourceStream = assembly.GetManifestResourceStream("MultiplyChannels.NewDevice.xml");
            using (var reader = new StreamReader(resourceStream, Encoding.UTF8)) {
                lXmlFile = reader.ReadToEnd();
            }
            lXmlFile = lXmlFile.Replace("%ApplicationName%", opts.ApplicationName);
            lXmlFile = lXmlFile.Replace("%ApplicationNumber%", opts.ApplicationNumber.ToString());
            lXmlFile = lXmlFile.Replace("%ApplicationVersion%", opts.ApplicationVersion.ToString());
            lXmlFile = lXmlFile.Replace("%HardwareName%", opts.HardwareName);
            lXmlFile = lXmlFile.Replace("%HardwareVersion%", opts.HardwareVersion.ToString());
            lXmlFile = lXmlFile.Replace("%SerialNumber%", opts.SerialNumber);
            lXmlFile = lXmlFile.Replace("%OrderNumber%", opts.OrderNumber);
            lXmlFile = lXmlFile.Replace("%ProductName%", opts.ProductName);
            lXmlFile = lXmlFile.Replace("%MaskVersion%", opts.MaskVersion);
            lXmlFile = lXmlFile.Replace("%MediumTypes%", opts.MediumTypes);
            Console.WriteLine("Creating xml file {0}", opts.XmlFileName);
            File.WriteAllText(opts.XmlFileName, lXmlFile);
            return VerbCreate(opts);
        }

        static private int VerbCreate(CreateOptions opts) {
            string lHeaderFileName = Path.ChangeExtension(opts.XmlFileName, "h");
            if (opts.HeaderFileName != "") lHeaderFileName = opts.HeaderFileName;
            Console.WriteLine("Processing xml file {0}", opts.XmlFileName);
            ProcessInclude lResult = ProcessInclude.Factory(opts.XmlFileName, lHeaderFileName, opts.Prefix);
            lResult.Expand();
            // We restore the original namespace in File
            lResult.SetNamespace();
            XmlDocument lXml = lResult.GetDocument();
            bool lSuccess = ProcessSanityChecks(lXml);
            string lTempXmlFileName = Path.GetTempFileName();
            File.Delete(lTempXmlFileName);
            if (opts.Debug) lTempXmlFileName = opts.XmlFileName;
            lTempXmlFileName = Path.ChangeExtension(lTempXmlFileName, "debug.xml");
            if (opts.Debug) Console.WriteLine("Writing debug file to {0}", lTempXmlFileName);
            lXml.Save(lTempXmlFileName);
            Console.WriteLine("Writing header file to {0}", lHeaderFileName);
            File.WriteAllText(lHeaderFileName, lResult.HeaderGenerated);
            string lOutputFileName = Path.ChangeExtension(opts.OutputFile, "knxprod");
            if (opts.OutputFile == "") lOutputFileName = Path.ChangeExtension(opts.XmlFileName, "knxprod");
            if (lSuccess) {
                string lEtsPath = FindEtsPath(lTempXmlFileName);
                ExportKnxprod(lEtsPath, lTempXmlFileName, lOutputFileName);
            } else {
                Console.WriteLine("--> Skipping creation of {0} due to check errors! <--", lOutputFileName);
            }
            if (!opts.Debug) File.Delete(lTempXmlFileName);
            return 0;
        }

        static private int VerbCheck(CheckOptions opts) {
            string lFileName = Path.ChangeExtension(opts.XmlFileName, "xml");
            Console.WriteLine("Reading and resolving xml file {0}", lFileName);
            ProcessInclude lResult = ProcessInclude.Factory(opts.XmlFileName, "", "");
            lResult.LoadAdvanced(lFileName);
            return ProcessSanityChecks(lResult.GetDocument()) ? 0 : 1;
        }

        static private int VerbKnxprod(KnxprodOptions opts) {
            string lOutputFileName = Path.ChangeExtension(opts.OutputFile, "knxprod");
            if (opts.OutputFile == "") lOutputFileName = Path.ChangeExtension(opts.XmlFileName, "knxprod");
            Console.WriteLine("Reading xml file {0} writing to {1}", opts.XmlFileName, lOutputFileName);
            string lEtsPath = FindEtsPath(opts.XmlFileName);
            ExportKnxprod(lEtsPath, opts.XmlFileName, lOutputFileName);
            return 0;
        }
    }
}
