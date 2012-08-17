namespace RequirementsCompiler
{
    using System;
    using System.Reflection;
    using System.Xml;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class Program
    {
        // http://msdn.microsoft.com/en-us/library/cc850833.aspx
        public static void Main(string[] args)
        {
            new DocumentBuilder("modules.xml").Build();

            Console.WriteLine("Hello");
            Console.ReadKey();
        }        
    }
}
