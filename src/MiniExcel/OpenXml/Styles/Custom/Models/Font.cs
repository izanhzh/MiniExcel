using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml.Styles.Custom.Models
{
    public class Font
    {
        public string Name { get; set; }

        public double Size { get; set; }

        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public Underline Underline { get; set; }

        public bool Strike { get; set; }

        public string Color { get; set; }

        public VerticalAlign VertAlign { get; set; }

        public string Scheme { get; set; }

        public FontFamily Family { get; set; }

        public Font(
            string name,
            double size,
            bool bold = false,
            bool italic = false,
            Underline underline = Underline.None,
            bool strike = false,
            string color = null,
            VerticalAlign vertAlign = VerticalAlign.Baseline,
            string scheme = null,
            FontFamily family = FontFamily.Automatic)
        {
            Name = name;
            Size = size;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strike = strike;
            Color = color;
            VertAlign = vertAlign;
            Scheme = scheme;
            Family = family;
        }

        internal void WriteToXml(SheetStyleBuildContext context)
        {
            context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "font", context.OldXmlReader.NamespaceURI);

            if (!string.IsNullOrEmpty(Name))
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "name", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("val", Name);
                context.NewXmlWriter.WriteEndElement();
            }

            context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "sz", context.OldXmlReader.NamespaceURI);
            context.NewXmlWriter.WriteAttributeString("val", Size.ToString());
            context.NewXmlWriter.WriteEndElement();

            if (Bold)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "b", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteEndElement();
            }

            if (Italic)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "i", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteEndElement();
            }

            if (Underline != Underline.None)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "u", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("val", Underline.ToString().ToLower());
                context.NewXmlWriter.WriteEndElement();
            }

            if (Strike)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "strike", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteEndElement();
            }

            if (!string.IsNullOrEmpty(Color))
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("rgb", Color);
                context.NewXmlWriter.WriteEndElement();
            }

            if (VertAlign != VerticalAlign.Baseline)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "vertAlign", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("val", VertAlign.ToString().ToLower());
                context.NewXmlWriter.WriteEndElement();
            }

            if (!string.IsNullOrEmpty(Scheme))
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "scheme", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("val", Scheme);
                context.NewXmlWriter.WriteEndElement();
            }

            context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "family", context.OldXmlReader.NamespaceURI);
            context.NewXmlWriter.WriteAttributeString("val", ((int)Family).ToString());
            context.NewXmlWriter.WriteEndElement();

            context.NewXmlWriter.WriteEndElement();
        }

        internal async Task WriteToXmlAsync(SheetStyleBuildContext context)
        {
            await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "font", context.OldXmlReader.NamespaceURI);

            if (!string.IsNullOrEmpty(Name))
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "name", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, Name);
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "sz", context.OldXmlReader.NamespaceURI);
            await context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, Size.ToString());
            await context.NewXmlWriter.WriteEndElementAsync();

            if (Bold)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "b", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (Italic)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "i", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (Underline != Underline.None)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "u", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, Underline.ToString().ToLower());
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (Strike)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "strike", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (!string.IsNullOrEmpty(Color))
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, Color);
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (VertAlign != VerticalAlign.Baseline)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "vertAlign", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, VertAlign.ToString().ToLower());
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (!string.IsNullOrEmpty(Scheme))
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "scheme", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, Scheme);
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "family", context.OldXmlReader.NamespaceURI);
            await context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, ((int)Family).ToString());
            await context.NewXmlWriter.WriteEndElementAsync();

            await context.NewXmlWriter.WriteEndElementAsync();
        }
    }
}
