using System;
using System.IO;
using GemBox.Document;

namespace WordToPdf.GemBoxDocument
{
    public sealed class PdfParser : BaseParser
    {
        public PdfParser(string input) : base(input)
        {
        }

        public override void ConvertTo(string output, string password = "")
        {
            LicensePackaging.Setup();

            var docModel = DocumentModel.Load(Input);

            if (!string.IsNullOrWhiteSpace(password))
            {
                docModel.Save(output, new PdfSaveOptions
                {
                    Permissions = PdfPermissions.All,
                    PermissionsPassword = password,
                    DocumentOpenPassword = password
                });
            }
            else
            {
                docModel.Save(output);
            }

        }
    }
}