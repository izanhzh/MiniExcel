using System;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles.Custom.Models
{
    public class CellStyleXf
    {
        public int? CustomNumberFormatIndex { get; set; }

        public int? NumberFormatId { get; set; }

        public int CustomFontIndex { get; set; }

        public int CustomFillIndex { get; set; }

        public int CustomBorderIndex { get; set; }

        public bool ApplyNumberFormat { get; set; }

        public bool ApplyFont { get; set; }

        public bool ApplyFill { get; set; }

        public bool ApplyBorder { get; set; }

        public bool ApplyAlignment { get; set; }

        public bool ApplyProtection { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; }

        public VerticalAlignment VerticalAlignment { get; set; }

        public bool TextRotation { get; set; }

        public bool WrapText { get; set; }

        public bool ShrinkToFit { get; set; }

        public int Indent { get; set; }

        public bool Locked { get; set; }

        public bool Hidden { get; set; }

        public CellStyleXf(
            int? customNumberFormatIndex = null,
            int? numberFormatId = null,
            int customFontIndex = 0,
            int customFillIndex = 0,
            int customBorderIndex = 0,
            bool applyNumberFormat = false,
            bool applyFont = false,
            bool applyFill = false,
            bool applyBorder = false,
            bool applyAlignment = false,
            bool applyProtection = false,
            HorizontalAlignment horizontalAlignment = HorizontalAlignment.General,
            VerticalAlignment verticalAlignment = VerticalAlignment.Bottom,
            bool textRotation = false,
            bool wrapText = false,
            bool shrinkToFit = false,
            int indent = 0,
            bool locked = true,
            bool hidden = false)
        {
            if ((CustomNumberFormatIndex.HasValue && NumberFormatId.HasValue) || (!CustomNumberFormatIndex.HasValue && !NumberFormatId.HasValue))
            {
                throw new ArgumentException("Either CustomNumberFormatIndex or NumberFormatId must be provided");
            }
            CustomNumberFormatIndex = customNumberFormatIndex;
            NumberFormatId = numberFormatId;
            CustomFontIndex = customFontIndex;
            CustomFillIndex = customFillIndex;
            CustomBorderIndex = customBorderIndex;
            ApplyNumberFormat = applyNumberFormat;
            ApplyFont = applyFont;
            ApplyFill = applyFill;
            ApplyBorder = applyBorder;
            ApplyAlignment = applyAlignment;
            ApplyProtection = applyProtection;
            HorizontalAlignment = horizontalAlignment;
            VerticalAlignment = verticalAlignment;
            TextRotation = textRotation;
            WrapText = wrapText;
            ShrinkToFit = shrinkToFit;
            Indent = indent;
            Locked = locked;
            Hidden = hidden;
        }

        internal void WriteToXml(XmlWriter writer, string prefix, string namespaceUri, CustomStyleProviderBase customStyleProvider)
        {
            writer.WriteStartElement(prefix, "xf", namespaceUri);
            writer.WriteAttributeString("numFmtId", NumberFormatId.ToString());
            //writer.WriteAttributeString("fontId", FontId.ToString());
            //writer.WriteAttributeString("fillId", FillId.ToString());
            //writer.WriteAttributeString("borderId", BorderId.ToString());
            writer.WriteAttributeString("applyNumberFormat", ApplyNumberFormat.ToString().ToLower());
            writer.WriteAttributeString("applyFont", ApplyFont.ToString().ToLower());
            writer.WriteAttributeString("applyFill", ApplyFill.ToString().ToLower());
            writer.WriteAttributeString("applyBorder", ApplyBorder.ToString().ToLower());
            writer.WriteAttributeString("applyAlignment", ApplyAlignment.ToString().ToLower());
            writer.WriteAttributeString("applyProtection", ApplyProtection.ToString().ToLower());

            if (ApplyAlignment)
            {
                writer.WriteStartElement(prefix, "alignment", namespaceUri);
                writer.WriteAttributeString("horizontal", HorizontalAlignment.ToString().ToLower());
                writer.WriteAttributeString("vertical", VerticalAlignment.ToString().ToLower());
                if (TextRotation)
                {
                    writer.WriteAttributeString("textRotation", "90");
                }
                writer.WriteAttributeString("wrapText", WrapText.ToString().ToLower());
                writer.WriteAttributeString("shrinkToFit", ShrinkToFit.ToString().ToLower());
                writer.WriteAttributeString("indent", Indent.ToString());
                writer.WriteEndElement();
            }

            if (ApplyProtection)
            {
                writer.WriteStartElement(prefix, "protection", namespaceUri);
                writer.WriteAttributeString("locked", Locked.ToString().ToLower());
                writer.WriteAttributeString("hidden", Hidden.ToString().ToLower());
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        internal async Task WriteToXmlAsync(XmlWriter writer, string prefix, string namespaceUri, CustomStyleProviderBase customStyleProvider)
        {
            await writer.WriteStartElementAsync(prefix, "xf", namespaceUri);
            await writer.WriteAttributeStringAsync(null, "numFmtId", null, NumberFormatId.ToString());
            //await writer.WriteAttributeStringAsync(null, "fontId", null, FontId.ToString());
            //await writer.WriteAttributeStringAsync(null, "fillId", null, FillId.ToString());
            //await writer.WriteAttributeStringAsync(null, "borderId", null, BorderId.ToString());
            await writer.WriteAttributeStringAsync(null, "applyNumberFormat", null, ApplyNumberFormat.ToString().ToLower());
            await writer.WriteAttributeStringAsync(null, "applyFont", null, ApplyFont.ToString().ToLower());
            await writer.WriteAttributeStringAsync(null, "applyFill", null, ApplyFill.ToString().ToLower());
            await writer.WriteAttributeStringAsync(null, "applyBorder", null, ApplyBorder.ToString().ToLower());
            await writer.WriteAttributeStringAsync(null, "applyAlignment", null, ApplyAlignment.ToString().ToLower());
            await writer.WriteAttributeStringAsync(null, "applyProtection", null, ApplyProtection.ToString().ToLower());

            if (ApplyAlignment)
            {
                await writer.WriteStartElementAsync(prefix, "alignment", namespaceUri);
                await writer.WriteAttributeStringAsync(null, "horizontal", null, HorizontalAlignment.ToString().ToLower());
                await writer.WriteAttributeStringAsync(null, "vertical", null, VerticalAlignment.ToString().ToLower());
                if (TextRotation)
                {
                    await writer.WriteAttributeStringAsync(null, "textRotation", null, "90");
                }
                await writer.WriteAttributeStringAsync(null, "wrapText", null, WrapText.ToString().ToLower());
                await writer.WriteAttributeStringAsync(null, "shrinkToFit", null, ShrinkToFit.ToString().ToLower());
                await writer.WriteAttributeStringAsync(null, "indent", null, Indent.ToString());
                await writer.WriteEndElementAsync();
            }

            if (ApplyProtection)
            {
                await writer.WriteStartElementAsync(prefix, "protection", namespaceUri);
                await writer.WriteAttributeStringAsync(null, "locked", null, Locked.ToString().ToLower());
                await writer.WriteAttributeStringAsync(null, "hidden", null, Hidden.ToString().ToLower());
                await writer.WriteEndElementAsync();
            }

            await writer.WriteEndElementAsync();
        }
    }
}
