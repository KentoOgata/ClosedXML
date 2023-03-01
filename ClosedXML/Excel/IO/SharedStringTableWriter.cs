using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.IO;
using ClosedXML.Extensions;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;
using System.Xml.Linq;

namespace ClosedXML.Excel.IO
{
    internal class SharedStringTableWriter
    {
        internal static void GenerateSharedStringTablePartContent(XLWorkbook workbook, SharedStringTablePart sharedStringTablePart,
            SaveContext context)
        {
            // Call all table headers to make sure their names are filled
            workbook.Worksheets.ForEach(w => w.Tables.ForEach(t => _ = ((XLTable)t).FieldNames.Count));

            var stringId = 0;

            var newStrings = new Dictionary<String, Int32>();
            var newRichStrings = new Dictionary<IXLRichText, Int32>();

            static bool HasSharedString(XLCell c)
            {
                if (c.DataType == XLDataType.Text && c.ShareString)
                    return c.StyleValue.IncludeQuotePrefix || String.IsNullOrWhiteSpace(c.FormulaA1) && c.GetText().Length > 0;
                else
                    return false;
            }

            var settings = new XmlWriterSettings
            {
                CloseOutput = true,
                Encoding = XLHelper.NoBomUTF8
            };
            var partStream = sharedStringTablePart.GetStream(FileMode.Create);
            using var xml = XmlWriter.Create(partStream, settings);

            var sst = new XElement(XMain2006SsNs + "sst");

            int countOfTextStringsInWorkbook = 0;

            foreach (var c in workbook.Worksheets.Cast<XLWorksheet>().SelectMany(w => w.Internals.CellsCollection.GetCells(HasSharedString)))
            {
                countOfTextStringsInWorkbook++;
                if (c.HasRichText)
                {
                    if (newRichStrings.TryGetValue(c.GetRichText(), out int id))
                        c.SharedStringId = id;
                    else
                    {
                        var si = new XElement(XMain2006SsNs + "si");
                        si.AddRichTextElements(c, context);

                        sst.Add(si);

                        newRichStrings.Add(c.GetRichText(), stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
                else
                {
                    var value = c.Value.GetText();
                    if (newStrings.TryGetValue(value, out int id))
                        c.SharedStringId = id;
                    else
                    {
                        var t = new XElement(XMain2006SsNs + "t");
                        var sharedString = value;
                        if (!sharedString.Trim().Equals(sharedString))
                            t.AddPreserveSpaceAttr();

                        t.SetValue(XmlEncoder.EncodeString(sharedString));

                        var si = new XElement(XMain2006SsNs + "si", t);
                        sst.Add(si);

                        newStrings.Add(value, stringId);
                        c.SharedStringId = stringId;

                        stringId++;
                    }
                }
            }

            sst.Add(new XAttribute("count", countOfTextStringsInWorkbook));
            sst.Add(new XAttribute("uniqueCount", sst.Elements().Count()));

            var sharedStrings = new XDocument(new XDeclaration(version: "1.0", encoding: "utf-8", standalone: "yes"), sst);

            sharedStrings.WriteTo(xml);
            xml.Close();
        }
    }
}
