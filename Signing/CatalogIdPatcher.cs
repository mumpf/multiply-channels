using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;

namespace MultiplyChannels.Signing
{
    class CatalogIdPatcher
    {
        public CatalogIdPatcher(
            FileInfo catalogFile,
            IDictionary<string, string> hardware2ProgramIdMapping,
            string basePath,
            int namespaceVersion)
        {
            Assembly asm = Assembly.LoadFrom(Path.Combine(basePath, "Knx.Ets.XmlSigning.dll"));

            if(basePath.Contains("ETS6") || basePath.Contains("6.0")) {
                _instance = Activator.CreateInstance(asm.GetType("Knx.Ets.XmlSigning.CatalogIdPatcher"), catalogFile, hardware2ProgramIdMapping, null);
                _type = asm.GetType("Knx.Ets.XmlSigning.CatalogIdPatcher");
            } else {
                _instance = Activator.CreateInstance(asm.GetType("Knx.Ets.XmlSigning.CatalogIdPatcher"), catalogFile, hardware2ProgramIdMapping);
                _type = asm.GetType("Knx.Ets.XmlSigning.CatalogIdPatcher");
            }
        }

        public void Patch()
        {
            _type.GetMethod("Patch", BindingFlags.Instance | BindingFlags.Public).Invoke(_instance, null);
        }

        private readonly object _instance;
        private readonly Type _type;
    }
}