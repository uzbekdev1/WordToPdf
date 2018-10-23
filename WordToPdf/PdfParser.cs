using System;
using System.IO;

namespace WordToPdf
{
    public static class PdfParser
    {
        public static void Convert(string input, string output)
        {
            var args = "-f \"" + input + "\" -O \"" + output + "\" -T wdFormatPDF";
            var pathExe = Path.Combine(Environment.CurrentDirectory, "Converter", "docto.exe");

            ProcessHelper.Run(pathExe, args, true);
        }
    }
}