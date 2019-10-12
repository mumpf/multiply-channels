using System.Xml;
using System.IO;

namespace MultiplyChannels {
    public static class ExtensionMethods {
        public static string GetNamedItemValueOrEmpty(this XmlAttributeCollection iAttibutes, string iName) {
            string lResult = "";
            XmlNode lAttribute = iAttibutes.GetNamedItem(iName);
            if (lAttribute != null) lResult = lAttribute.Value;
            return lResult;
        }

        public static string NodeName(this XmlNode iNode) {
            string lResult = "";
            XmlNode lAttribute = iNode.Attributes.GetNamedItem("Name");
            if (lAttribute != null) lResult = lAttribute.Value.ToString();
            return lResult;
        }


    }
}

