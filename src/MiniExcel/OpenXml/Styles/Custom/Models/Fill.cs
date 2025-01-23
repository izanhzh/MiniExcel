using System;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles.Custom.Models
{
    public class Fill
    {
        public string SolidColor { get; set; }

        public string PatternType { get; set; }

        public string PatternForegroundColor { get; set; }

        public string PatternBackgroundColor { get; set; }

        public Fill(string solidColor = null, string patternType = null, string patternForegroundColor = null, string patternBackgroundColor = null)
        {
            if (string.IsNullOrEmpty(solidColor) && string.IsNullOrEmpty(patternType))
            {
                throw new ArgumentException("Either SolidColor or PatternType must be provided.");
            }

            if (!string.IsNullOrEmpty(patternType) && string.IsNullOrEmpty(patternForegroundColor))
            {
                throw new ArgumentException("PatternForegroundColor cannot be null or empty when PatternType is specified.");
            }

            SolidColor = solidColor;
            PatternType = patternType;
            PatternForegroundColor = patternForegroundColor;
            PatternBackgroundColor = patternBackgroundColor;
        }

        internal void WriteToXml(XmlWriter writer, string prefix, string namespaceUri)
        {
            writer.WriteStartElement(prefix, "fill", namespaceUri);

            if (!string.IsNullOrEmpty(SolidColor))
            {
                writer.WriteStartElement(prefix, "solidFill", namespaceUri);
                writer.WriteStartElement(prefix, "fgColor", namespaceUri);
                writer.WriteAttributeString("rgb", SolidColor);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            if (!string.IsNullOrEmpty(PatternType))
            {
                writer.WriteStartElement(prefix, "patternFill", namespaceUri);
                writer.WriteAttributeString("patternType", PatternType);

                if (!string.IsNullOrEmpty(PatternForegroundColor))
                {
                    writer.WriteStartElement(prefix, "fgColor", namespaceUri);
                    writer.WriteAttributeString("rgb", PatternForegroundColor);
                    writer.WriteEndElement();
                }

                if (!string.IsNullOrEmpty(PatternBackgroundColor))
                {
                    writer.WriteStartElement(prefix, "bgColor", namespaceUri);
                    writer.WriteAttributeString("rgb", PatternBackgroundColor);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        internal async Task WriteToXmlAsync(XmlWriter writer, string prefix, string namespaceUri)
        {
            await writer.WriteStartElementAsync(prefix, "fill", namespaceUri);

            if (!string.IsNullOrEmpty(SolidColor))
            {
                await writer.WriteStartElementAsync(prefix, "solidFill", namespaceUri);
                await writer.WriteStartElementAsync(prefix, "fgColor", namespaceUri);
                await writer.WriteAttributeStringAsync(null, "rgb", null, SolidColor);
                await writer.WriteEndElementAsync();
                await writer.WriteEndElementAsync();
            }

            if (!string.IsNullOrEmpty(PatternType))
            {
                await writer.WriteStartElementAsync(prefix, "patternFill", namespaceUri);
                await writer.WriteAttributeStringAsync(null, "patternType", null, PatternType);

                if (!string.IsNullOrEmpty(PatternForegroundColor))
                {
                    await writer.WriteStartElementAsync(prefix, "fgColor", namespaceUri);
                    await writer.WriteAttributeStringAsync(null, "rgb", null, PatternForegroundColor);
                    await writer.WriteEndElementAsync();
                }

                if (!string.IsNullOrEmpty(PatternBackgroundColor))
                {
                    await writer.WriteStartElementAsync(prefix, "bgColor", namespaceUri);
                    await writer.WriteAttributeStringAsync(null, "rgb", null, PatternBackgroundColor);
                    await writer.WriteEndElementAsync();
                }

                await writer.WriteEndElementAsync();
            }

            await writer.WriteEndElementAsync();
        }
    }
}
