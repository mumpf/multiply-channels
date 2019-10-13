﻿using System;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Text;
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
                string lXmlContent = File.ReadLines(iXmlFilename).Take(2).Aggregate((s1, s2) => s1 + s2 );
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

        static bool ProcessSanityChecks(XmlNode iTargetNode) {

            Console.WriteLine("Sanity checks... ");
            bool lFail = false;

            Console.Write("- Id-Uniqueness...");
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
            if (!lFailPart) Console.WriteLine(" finished");
            lFail = lFail || lFailPart;

            lFailPart = false;
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
            if (!lFailPart) Console.WriteLine(" finished");
            lFail = lFail || lFailPart;

            lFailPart = false;
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
            if (!lFailPart) Console.WriteLine(" finished");
            lFail = lFail || lFailPart;

            lFailPart = false;
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
            if (!lFailPart) Console.WriteLine(" finished");
            lFail = lFail || lFailPart;

            lFailPart = false;
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
            if (!lFailPart) Console.WriteLine(" finished");
            lFail = lFail || lFailPart;

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
            [Value(0, MetaName = "xml input file name", HelpText = "Xml input file name")]
            public string XmlFileName {
                get { return mXmlFileName; }
                set { mXmlFileName = Path.ChangeExtension(value, "xml"); }
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
            // We restore the original namespace in File
            lResult.SetNamespace();
            XmlDocument lXml = lResult.GetDocument();
            bool lSuccess = ProcessSanityChecks(lXml);
            string lTempXmlFileName = Path.ChangeExtension(opts.XmlFileName, "out.xml");
            Console.WriteLine("Writing intermediate file to {0}", lTempXmlFileName);
            lXml.Save(lTempXmlFileName);
            Console.WriteLine("Writing header file to {0}", lHeaderFileName);
            File.WriteAllText(lHeaderFileName, lResult.HeaderGenerated);
            string lOutputFileName = Path.ChangeExtension(opts.OutputFile, "knxprod");
            if (opts.OutputFile == "") lOutputFileName = Path.ChangeExtension(opts.XmlFileName, "knxprod");
            if (lSuccess) {
                string lEtsPath = FindEtsPath(lTempXmlFileName);
                ExportKnxprod(lEtsPath, lTempXmlFileName, lOutputFileName);
            }
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