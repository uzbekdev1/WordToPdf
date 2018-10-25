using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace WordToPdf
{
    public static class WordReplacer
    {
        /// <summary>
        /// Key is include #{KEY}#=Value
        /// </summary>
        /// <param name="path"></param>
        /// <param name="keyValues"></param>
        public static void DoReplace(string path, IDictionary<string, string> keyValues)
        {
            using (var doc = WordprocessingDocument.Open(path, true))
            {
                foreach (var keyValue in keyValues)
                {
                    TextReplacer.SearchAndReplace(doc, "#" + keyValue.Key + "#", keyValue.Value, true);
                }
            }
        }
    }
}