using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using WordToPdf;

namespace Demo
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
            var keyValues = new Dictionary<string, string>
            {
                {"SELLER_ORG", "ORG1"},
                {"BUYER_ORG", "ORG2"},
                {"DEAL_NUMBER", "1"},
                {"DEAL_DATE", DateTime.Today.ToShortDateString()}
            };
            WordReplacer.DoReplace(pathDest, keyValues);

            //converting process
            var pathConvert = Path.Combine(ReflectionProperty.AssemblyFolder, "Upload", "T1.pdf");
            //WithOfficeDoParsing(pathDest, pathConvert);
            WithoutOfficeDoParsing(pathDest, pathConvert);

            //run in process result
            Process.Start(pathConvert);
        }

        private static void WithOfficeDoParsing(string input, string output)
        {
            var pdfParser = new WordToPdf.docto.PdfParser(input);
            pdfParser.ConvertTo(output);
        }

        private static void WithoutOfficeDoParsing(string input, string output)
        {
            var pdfParser = new WordToPdf.GemBoxDocument.PdfParser(input);
            pdfParser.ConvertTo(output);
        }

    }
}