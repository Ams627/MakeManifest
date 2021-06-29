using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace MakeManifest
{
    class TypeLibParser
    {
        private readonly string _filename;

        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern ITypeLib LoadTypeLib([In, MarshalAs(UnmanagedType.LPWStr)] string typelib);

        public TypeLibParser(string filename)
        {
            this._filename = filename;
        }

        public XDocument ParseTypeLib()
        {
            ITypeLib typeLib = LoadTypeLib(_filename);

            typeLib.GetLibAttr(out IntPtr ipLibAtt);

            var typeLibAttr = (System.Runtime.InteropServices.ComTypes.TYPELIBATTR)
                Marshal.PtrToStructure(ipLibAtt, typeof(System.Runtime.InteropServices.ComTypes.TYPELIBATTR));

            XNamespace ns = "urn:schemas-microsoft-com:asm.v1";
            var document = new XDocument(
                new XElement(ns + "assembly",
                new XElement(ns + "assemblyIdentity", new XAttribute("type", "win32"), new XAttribute("name", _filename),
                new XAttribute("version", "1.0.0.0"))));


            string fileNameOnly = Path.GetFileNameWithoutExtension(_filename);
            var comClassElems = new List<XElement>();

            for (int i = 0; i < typeLib.GetTypeInfoCount(); i++)
            {
                typeLib.GetTypeInfo(i, out ITypeInfo typeInfo);

                // figure out what guids, typekind, and names of the thing we're dealing with
                typeInfo.GetTypeAttr(out IntPtr ipTypeAttr);

                // unmarshal the pointer into a structure into something we can read
                var typeattr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)
                    Marshal.PtrToStructure(ipTypeAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

                System.Runtime.InteropServices.ComTypes.TYPEKIND typeKind = typeattr.typekind;
                Guid typeId = typeattr.guid;

                //get the name of the type
                typeLib.GetDocumentation(i, out string name, out string docString, out int dwHelpContext, out string strHelpFile);


                if (typeKind == System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_COCLASS)
                {
                    var elem = new XElement(ns + "comClass",
                        new XAttribute("clsid", typeId),
                        new XAttribute("description", docString),
                        new XAttribute("progid", $"{fileNameOnly}.{name}")
                        );
                    comClassElems.Add(elem);
                }
            }

            document.Root.Add(new XElement(ns + "file", new XAttribute("name", _filename), comClassElems));

            return document;
        }
    }
}
