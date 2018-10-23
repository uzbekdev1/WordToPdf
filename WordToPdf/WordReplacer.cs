using System;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace WordToPdf
{
    public static class WordReplacer
    {
        public static void DoReplace(string path)
        {
            var orgSeller = "SELLER_ORG";
            var orgBuyer = "BUYER_ORG";
            var dealNum = "DEAL_NUMBER";
            var dealDate = "DEAL_DATE";

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                TextReplacer.SearchAndReplace(doc, "#" + orgSeller + "#", "ORG1", true);
                TextReplacer.SearchAndReplace(doc, "#" + orgBuyer + "#", "ORG2", true);
                TextReplacer.SearchAndReplace(doc, "#" + dealNum + "#", "1", true);
                TextReplacer.SearchAndReplace(doc, "#" + dealDate + "#", DateTime.Today.ToShortDateString(), true);
            }
        }
    }
}