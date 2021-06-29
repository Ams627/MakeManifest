using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Threading.Tasks;

namespace MakeManifest
{
    class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                foreach (var arg in args)
                {
                    var parser = new TypeLibParser(arg);
                    var doc = parser.ParseTypeLib();
                    Console.WriteLine(doc);
                }
            }
            catch (Exception ex)
            {
                var fullname = System.Reflection.Assembly.GetEntryAssembly().Location;
                var progname = Path.GetFileNameWithoutExtension(fullname);
                Console.Error.WriteLine($"{progname} Error: {ex.Message}");
            }

        }
    }
}
