using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles.Custom.Models
{
    public class NumFmt
    {
        public int NumFmtId { get; set; }

        public string FormatCode { get; set; }

        public NumFmt(int numFmtId, string formatCode)
        {
            NumFmtId = numFmtId;
            FormatCode = formatCode;
        }

        internal void WriteToXml(SheetStyleBuildContext context)
        {
            context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "numFmt", context.OldXmlReader.NamespaceURI);
            context.NewXmlWriter.WriteAttributeString("numFmtId", NumFmtId.ToString());
            context.NewXmlWriter.WriteAttributeString("formatCode", FormatCode);
            context.NewXmlWriter.WriteEndElement();
        }

        internal async Task WriteToXmlAsync(SheetStyleBuildContext context)
        {
            await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "numFmt", context.OldXmlReader.NamespaceURI);
            await context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, NumFmtId.ToString());
            await context.NewXmlWriter.WriteAttributeStringAsync(null, "formatCode", null, FormatCode);
            await context.NewXmlWriter.WriteEndElementAsync();
        }
    }
}
