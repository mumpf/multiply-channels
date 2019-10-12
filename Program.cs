using System;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Text;
using CommandLine;
// using Knx.Ets.Xml.ObjectModel;

//using Knx.Ets.Converter.ConverterEngine;

namespace MultiplyChannels {
    class Program {

        private const string gtoolName = "KNX MT";
        private const string gtoolVersion = "5.1.255.16695";
        //installation path of a valid ETS instance (only ETS4 or ETS5 supported)
        private static string gPathETS = @"C:\Program Files (x86)\ETS";


        static bool ProcessSanityChecks(XmlNode iTargetNode, EtsVersions iVersion) {

            Console.WriteLine("Sanity checks... ");
            Console.Write("- Id-Uniqueness...");
            bool lFail = false;
            bool lFailPart = false;
            XmlNodeList lNodes = iTargetNode.SelectNodes("//*[@Id]");
            Dictionary<string, bool> lIds = new Dictionary<string, bool>();
            foreach (XmlNode lNode in lNodes) {
                string lId = lNode.Attributes.GetNamedItem("Id").Value;
                if (lIds.ContainsKey(lId)) {
                    if (!lFailPart) Console.WriteLine();
                    Console.WriteLine("  {0} is a duplicate Id in {1}", lId, lNode.NodeName());
                    lFailPart = true;
                } else {
                    lIds.Add(lId, true);
                }
            }
            if (!lFailPart) Console.WriteLine("finished");
            lFail = lFail || lFailPart;

            Console.Write("- RefId-Integrity...");
            lNodes = iTargetNode.SelectNodes("//*[@RefId]");
            foreach (XmlNode lNode in lNodes) {
                if (lNode.Name != "Manufacturer") {
                    string lRefId = lNode.Attributes.GetNamedItem("RefId").Value;
                    if (!lIds.ContainsKey(lRefId)) {
                        if (!lFailPart) Console.WriteLine();
                        Console.WriteLine("  {0} is referenced in {1} {2}, but not defined", lRefId, lNode.Name, lNode.NodeName());
                        lFailPart = true;
                    }
                }
            }
            if (!lFailPart) Console.WriteLine("finished");
            lFail = lFail || lFailPart;

            Console.Write("- ParamRefId-Integrity...");
            lNodes = iTargetNode.SelectNodes("//*[@ParamRefId]");
            foreach (XmlNode lNode in lNodes) {
                if (lNode.Name != "Manufacturer") {
                    string lParamRefId = lNode.Attributes.GetNamedItem("ParamRefId").Value;
                    if (!lIds.ContainsKey(lParamRefId)) {
                        if (!lFailPart) Console.WriteLine();
                        Console.WriteLine("  {0} is referenced in {1} {2}, but not defined", lParamRefId, lNode.Name, lNode.NodeName());
                        lFailPart = true;
                    }
                }
            }
            if (!lFailPart) Console.WriteLine("finished");
            lFail = lFail || lFailPart;

            Console.Write("- ParameterType-Integrity...");
            lNodes = iTargetNode.SelectNodes("//*[@ParameterType]");
            foreach (XmlNode lNode in lNodes) {
                string lParameterType = lNode.Attributes.GetNamedItem("ParameterType").Value;
                if (!lIds.ContainsKey(lParameterType)) {
                    if (!lFailPart) Console.WriteLine();
                    Console.WriteLine("  {0} is referenced in {1} {2}, but not defined", lParameterType, lNode.Name, lNode.NodeName());
                    lFailPart = true;
                }
            }
            if (!lFailPart) Console.WriteLine("finished");
            lFail = lFail || lFailPart;

            Console.Write("- Serial number...");
            lNodes = iTargetNode.SelectNodes("//*[@SerialNumber]");
            foreach (XmlNode lNode in lNodes) {
                string lSerialNumber = lNode.Attributes.GetNamedItem("SerialNumber").Value;
                if (lSerialNumber.Contains("-")) {
                    if (!lFailPart) Console.WriteLine();
                    Console.WriteLine("  Hardware.SerialNumber={0}, it contains a dash (-), this will cause problems in knxprod.", lSerialNumber);
                    lFailPart = true;
                }
            }
            if (!lFailPart) Console.WriteLine("finished");
            lFail = lFail || lFailPart;

            Console.Write("- xmlns...");
            if (iTargetNode.NodeType == XmlNodeType.Document) {
                string lSourceXmlns = ((XmlDocument)iTargetNode).DocumentElement.GetAttribute("xmlns");
                if (lSourceXmlns == "") {
                    lSourceXmlns = ((XmlDocument)iTargetNode).DocumentElement.GetAttribute("oldxmlns");
                }
                string lTargetXmlns = ProcessInclude.GetNamespace(iVersion);
                if (lSourceXmlns != lTargetXmlns) {
                    if (!lFailPart) Console.WriteLine();
                    Console.WriteLine("  xmlns is {0}, expected is {1}", lSourceXmlns, lTargetXmlns);
                    lFailPart = true;
                }
                if (!lFailPart) Console.WriteLine("finished");
                lFail = lFail || lFailPart;
            }
            return !lFail;
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

        private static void ExportKnxprod(string iXmlFileName, string iKnxprodFileName) {
            try {
                var files = new string[] { iXmlFileName };
                var asmPath = Path.Combine(gPathETS, "Knx.Ets.Converter.ConverterEngine.dll");
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
            [Value(0, MetaName = "xml input file name", HelpText = "Xml input file name")]
            public string XmlFileName {
                get { return mXmlFileName; }
                set { mXmlFileName = Path.ChangeExtension(value, "xml"); }
            }
            [Option('e', "EtsVersion", Default = "56", Required = false, HelpText = "use dll's of given version", MetaValue = "4,55,56,57")]
            public string EtsVersion { get; set; } = "56";

            public EtsVersions ParseVersion() {
                // for ETS5 conversion dll, we need to change the namespace in xml file
                string lEts = "!!! NO ETS !!!";
                string lPathEts = "";
                EtsVersions lResult;
                if (EtsVersion == "57") {
                    lResult = EtsVersions.Ets57;
                    lEts = "ETS-5.7";
                    lPathEts = "5";
                } else if (EtsVersion == "56") {
                    lResult = EtsVersions.Ets56;
                    lEts = "ETS-5.6";
                    lPathEts = "5";
                } else if (EtsVersion == "55") {
                    lResult = EtsVersions.Ets55;
                    lEts = "ETS-5.5";
                    lPathEts = "5";
                } else {
                    lResult = EtsVersions.Ets4;
                    lEts = "ETS-4";
                    lPathEts = "4";
                }
                if (gPathETS.EndsWith("ETS")) {
                    gPathETS += lPathEts;
                }
                Console.WriteLine("Using {0} for conversion", lEts);
                return lResult;
            }
        }

        [Verb("knxprod", HelpText = "Create knxprod file from given xml file")]
        class KnxprodOptions : EtsOptions {
            [Option('o', "Output", Required = false, HelpText = "output file name")]
            public string OutputFile { get; set; } = "";
        }

        [Verb("create", HelpText = "Process given xml file with all includes and create knxprod")]
        class CreateOptions : KnxprodOptions {
            [Option('h', "HeaderFileName", Required = false, HelpText = "Header file name")]
            public string HeaderFileName { get; set; } = "";
        }

        [Verb("check", HelpText = "execute sanity checks on given xml file")]
        class CheckOptions : EtsOptions {
        }

        static int Main(string[] args) {
            return CommandLine.Parser.Default.ParseArguments<CreateOptions, CheckOptions, KnxprodOptions>(args)
              .MapResult(
                (CreateOptions opts) => VerbCreate(opts),
                (CheckOptions opts) => VerbCheck(opts),
                (KnxprodOptions opts) => VerbKnxprod(opts),
                errs => 1);
        }

        static private int VerbCreate(CreateOptions opts) {
            string lHeaderFileName = Path.ChangeExtension(opts.XmlFileName, "h");
            if (opts.HeaderFileName != "") lHeaderFileName = opts.HeaderFileName;
            Console.WriteLine("Reading and processing xml file {0}", opts.XmlFileName);
            ProcessInclude lResult = ProcessInclude.Factory(opts.XmlFileName, lHeaderFileName, "");
            lResult.Expand();
            // for ETS5 conversion dll, we need to change the namespace in xml file
            lResult.SetNamespace(opts.ParseVersion());
            XmlDocument lXml = lResult.GetDocument();
            bool lSuccess = ProcessSanityChecks(lXml, opts.ParseVersion());
            string lTempXmlFileName = Path.ChangeExtension(opts.XmlFileName, "out.xml");
            Console.WriteLine("Writing intermediate file to {0}", lTempXmlFileName);
            lXml.Save(lTempXmlFileName);
            Console.WriteLine("Writing header file to {0}", lHeaderFileName);
            File.WriteAllText(lHeaderFileName, lResult.HeaderGenerated);
            string lOutputFileName = Path.ChangeExtension(opts.OutputFile, "knxprod");
            if (opts.OutputFile == "") lOutputFileName = Path.ChangeExtension(opts.XmlFileName, "knxprod");
            if (lSuccess) ExportKnxprod(lTempXmlFileName, lOutputFileName);
            return 0;
        }

        static private int VerbCheck(CheckOptions opts) {
            string lFileName = Path.ChangeExtension(opts.XmlFileName, "xml");
            Console.WriteLine("Reading and resolving xml file {0}", lFileName);
            ProcessInclude lResult = ProcessInclude.Factory(opts.XmlFileName, "", "");
            lResult.LoadAdvanced(lFileName);
            return ProcessSanityChecks(lResult.GetDocument(), opts.ParseVersion()) ? 0 : 1;
        }

        static private int VerbKnxprod(KnxprodOptions opts) {
            string lOutputFileName = Path.ChangeExtension(opts.OutputFile, "knxprod");
            if (opts.OutputFile == "") lOutputFileName = Path.ChangeExtension(opts.XmlFileName, "knxprod");
            Console.WriteLine("Reading xml file {0} writing to {1}", opts.XmlFileName, lOutputFileName);
            opts.ParseVersion();
            ExportKnxprod(opts.XmlFileName, lOutputFileName);
            return 0;
        }
    }
}
