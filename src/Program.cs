using System;
using System.Diagnostics;
using System.IO;
using WordToPdf;

namespace ConsoleApplication2
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //template file
            var pathSrc = Path.Combine(ReflectionProperty.AssemblyFolder, "Templates", "T1.docx");

            //result file
            var pathDest = Path.Combine(ReflectionProperty.AssemblyFolder, "Upload", "T1.docx");

            //copy template
            File.Copy(pathSrc, pathDest, true);

            //replace word
            WordReplacer.DoReplace(pathDest);

            //converting process
            var pathConvert = Path.Combine(ReflectionProperty.AssemblyFolder, "Upload", "T1.pdf");
            PdfParser.Convert(pathDest, pathConvert);

            //run in process result
            Process.Start(pathConvert);
        }
    }
}