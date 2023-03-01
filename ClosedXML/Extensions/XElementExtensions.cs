using System.Xml.Linq;
using ClosedXML.Excel;
using static ClosedXML.Excel.XLWorkbook;
using static ClosedXML.Excel.IO.OpenXmlConst;
using System;

namespace ClosedXML.Extensions
{
    internal static class XElementExtensions
    {
        internal static void AddRichTextElements(this XElement e, XLCell cell, SaveContext context)
        {
            var richText = cell.GetRichText();
            foreach (var rt in richText)
            {
                if (!string.IsNullOrEmpty(rt.Text))
                {
                    e.AddRichTextElement(rt);
                }
            }

            // Add phonetic element
            if (richText.HasPhonetics)
            {
                e.AddPhoneticsElements(richText.Phonetics, context);
            }
        }

        internal static void AddRichTextElement(this XElement e, XLRichString rt)
        {
            var rPr = new XElement(XMain2006SsNs + "rPr");

            if (rt.Bold)
                rPr.Add(new XElement(XMain2006SsNs + "b"));

            if (rt.Italic)
                rPr.Add(new XElement(XMain2006SsNs + "i"));

            if (rt.Strikethrough)
                rPr.Add(new XElement(XMain2006SsNs + "strike"));

            // Three attributes are not stored/written:
            // * outline - doesn't do anything and likely only works in Word.
            // * condense - legacy compatibility setting for macs
            // * extend - legacy compatibility setting for pre-xlsx Excels
            // None have sensible descriptions.

            if (rt.Shadow)
                rPr.Add(new XElement(XMain2006SsNs + "shadow"));

            if (rt.Underline != XLFontUnderlineValues.None)
                rPr.Add(new XElement(XMain2006SsNs + "u", new XAttribute("val", rt.Underline.ToOpenXmlString())));

            rPr.Add(new XElement(XMain2006SsNs + "vertAlign", new XAttribute("val", rt.VerticalAlignment.ToOpenXmlString())));
            rPr.Add(new XElement(XMain2006SsNs + "sz", new XAttribute("val", rt.FontSize)));
            var color = new XElement(XMain2006SsNs + "color");
            switch (rt.FontColor.ColorType)
            {
                case XLColorType.Color:
                    color.Add(new XAttribute("rgb", rt.FontColor.Color.ToHex()));
                    break;
                case XLColorType.Indexed:
                    color.Add(new XAttribute("indexed", rt.FontColor.Indexed));
                    break;
                case XLColorType.Theme:
                    color.Add(new XAttribute("theme", (int)rt.FontColor.ThemeColor));
                    if (rt.FontColor.ThemeTint != 0)
                        color.Add(new XAttribute("tint", rt.FontColor.ThemeTint));
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            rPr.Add(color);
            rPr.Add(new XElement(XMain2006SsNs + "rFont", new XAttribute("val", rt.FontName)));
            rPr.Add(new XElement(XMain2006SsNs + "family", new XAttribute("val", (Int32)rt.FontFamilyNumbering)));

            if (rt.FontCharSet != XLFontCharSet.Default)
                rPr.Add(new XElement(XMain2006SsNs + "charset", new XAttribute("val", (int)rt.FontCharSet)));

            if (rt.FontScheme != XLFontScheme.None)
                rPr.Add(new XElement(XMain2006SsNs + "scheme", new XAttribute("val", rt.FontScheme.ToOpenXml())));

            var t = new XElement(XMain2006SsNs + "t");
            if (rt.Text.PreserveSpaces())
                t.AddPreserveSpaceAttr();
            t.SetValue(rt.Text);

            var r = new XElement(XMain2006SsNs + "r", rPr, t);

            e.Add(r);
        }

        internal static void AddPhoneticsElements(this XElement e, IXLPhonetics phonetics, SaveContext saveContext)
        {
            foreach (var p in phonetics)
            {
                var rPh = new XElement(XMain2006SsNs + "rPh", new XAttribute("sb", p.Start), new XAttribute("eb", p.End));

                var t = new XElement(XMain2006SsNs + "t");
                if (p.Text.PreserveSpaces())
                    t.AddPreserveSpaceAttr();
                t.SetValue(p.Text);

                rPh.Add(t);

                e.Add(rPh);
            }

            var fontKey = XLFont.GenerateKey(phonetics);
            var f = XLFontValue.FromKey(ref fontKey);

            if (!saveContext.SharedFonts.TryGetValue(f, out FontInfo fi))
            {
                fi = new FontInfo { Font = f };
                saveContext.SharedFonts.Add(f, fi);
            }

            var phoneticPr = new XElement(XMain2006SsNs + "phoneticPr", new XAttribute("fontId", fi.FontId));

            if (phonetics.Alignment != XLPhoneticAlignment.Left)
                phoneticPr.Add(new XAttribute("alignment", phonetics.Alignment.ToOpenXmlString()));

            if (phonetics.Type != XLPhoneticType.FullWidthKatakana)
                phoneticPr.Add(new XAttribute("type", phonetics.Type.ToOpenXmlString()));

            e.Add(phoneticPr);
        }

        internal static void AddPreserveSpaceAttr(this XElement e)
        {
            e.Add(new XAttribute(XXml1998Ns + "space", "preserve"));
        }
    }
}
