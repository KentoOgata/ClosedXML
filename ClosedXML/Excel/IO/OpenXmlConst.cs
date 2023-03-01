using System;
using System.Xml.Linq;

namespace ClosedXML.Excel.IO
{
    /// <summary>
    /// Constants used across writers.
    /// </summary>
    internal static class OpenXmlConst
    {
        public const string Main2006SsNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        public static readonly XNamespace XMain2006SsNs = Main2006SsNs;

        public const string X14Ac2009SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

        public static readonly XNamespace XX14Ac2009SsNs = X14Ac2009SsNs;

        public const string Xml1998Ns = "http://www.w3.org/XML/1998/namespace";

        public static readonly XNamespace XXml1998Ns = Xml1998Ns;

        /// <summary>
        /// Valid and shorter than normal true.
        /// </summary>
        public static readonly String TrueValue = "1";

        /// <summary>
        /// Valid and shorter than normal false.
        /// </summary>
        public static readonly String FalseValue = "0";
    }
}
