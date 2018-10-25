using System;
using System.IO;

namespace WordToPdf.docto
{
    public class PdfParser : BaseParser
    {

        public PdfParser(string input) : base(input)
        {
        }

        public override void ConvertTo(string output, string password = "")
        {
            var args = "-f \"" + Input + "\" -O \"" + output + "\" -T wdFormatPDF";
            var pathExe = Path.Combine(Environment.CurrentDirectory, "Converter", "docto.exe");

            ProcessHelper.Run(pathExe, args, true);
        }

    }
}