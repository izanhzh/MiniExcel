using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles.Custom.Models
{
    public class Border
    {
        public BorderLine Left { get; set; }

        public BorderLine Right { get; set; }

        public BorderLine Top { get; set; }

        public BorderLine Bottom { get; set; }

        public BorderLine Diagonal { get; set; }

        public bool DiagonalUp { get; set; }

        public bool DiagonalDown { get; set; }

        public Border(
            BorderLine left = null,
            BorderLine right = null,
            BorderLine top = null,
            BorderLine bottom = null,
            BorderLine diagonal = null,
            bool diagonalUp = false,
            bool diagonalDown = false)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
            Diagonal = diagonal;
            DiagonalUp = diagonalUp;
            DiagonalDown = diagonalDown;
        }

        internal void WriteToXml(SheetStyleBuildContext context)
        {
            context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "border", context.OldXmlReader.NamespaceURI);

            context.NewXmlWriter.WriteAttributeString("diagonalUp", DiagonalUp ? "1" : "0");
            context.NewXmlWriter.WriteAttributeString("diagonalDown", DiagonalDown ? "1" : "0");

            if (Left != null && Left.Style != BorderStyle.None)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "left", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("style", Left.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Left.Color))
                {
                    context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    context.NewXmlWriter.WriteAttributeString("rgb", Left.Color);
                    context.NewXmlWriter.WriteEndElement();
                }
                context.NewXmlWriter.WriteEndElement();
            }

            if (Right != null && Right.Style != BorderStyle.None)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "right", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("style", Right.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Right.Color))
                {
                    context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    context.NewXmlWriter.WriteAttributeString("rgb", Right.Color);
                    context.NewXmlWriter.WriteEndElement();
                }
                context.NewXmlWriter.WriteEndElement();
            }

            if (Top != null && Top.Style != BorderStyle.None)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "top", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("style", Top.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Top.Color))
                {
                    context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    context.NewXmlWriter.WriteAttributeString("rgb", Top.Color);
                    context.NewXmlWriter.WriteEndElement();
                }
                context.NewXmlWriter.WriteEndElement();
            }

            if (Bottom != null && Bottom.Style != BorderStyle.None)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "bottom", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("style", Bottom.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Bottom.Color))
                {
                    context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    context.NewXmlWriter.WriteAttributeString("rgb", Bottom.Color);
                    context.NewXmlWriter.WriteEndElement();
                }
                context.NewXmlWriter.WriteEndElement();
            }

            if (Diagonal != null && Diagonal.Style != BorderStyle.None)
            {
                context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "diagonal", context.OldXmlReader.NamespaceURI);
                context.NewXmlWriter.WriteAttributeString("style", Diagonal.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Diagonal.Color))
                {
                    context.NewXmlWriter.WriteStartElement(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    context.NewXmlWriter.WriteAttributeString("rgb", Diagonal.Color);
                    context.NewXmlWriter.WriteEndElement();
                }
                context.NewXmlWriter.WriteEndElement();
            }

            context.NewXmlWriter.WriteEndElement();
        }

        internal async Task WriteToXmlAsync(SheetStyleBuildContext context)
        {
            await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "border", context.OldXmlReader.NamespaceURI);

            await context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalUp", null, DiagonalUp ? "1" : "0");
            await context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalDown", null, DiagonalDown ? "1" : "0");

            if (Left != null && Left.Style != BorderStyle.None)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "left", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, Left.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Left.Color))
                {
                    await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    await context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, Left.Color);
                    await context.NewXmlWriter.WriteEndElementAsync();
                }
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (Right != null && Right.Style != BorderStyle.None)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "right", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, Right.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Right.Color))
                {
                    await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    await context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, Right.Color);
                    await context.NewXmlWriter.WriteEndElementAsync();
                }
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (Top != null && Top.Style != BorderStyle.None)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "top", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, Top.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Top.Color))
                {
                    await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    await context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, Top.Color);
                    await context.NewXmlWriter.WriteEndElementAsync();
                }
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (Bottom != null && Bottom.Style != BorderStyle.None)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "bottom", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, Bottom.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Bottom.Color))
                {
                    await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    await context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, Bottom.Color);
                    await context.NewXmlWriter.WriteEndElementAsync();
                }
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            if (Diagonal != null && Diagonal.Style != BorderStyle.None)
            {
                await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "diagonal", context.OldXmlReader.NamespaceURI);
                await context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, Diagonal.Style.ToString().ToLower());
                if (!string.IsNullOrEmpty(Diagonal.Color))
                {
                    await context.NewXmlWriter.WriteStartElementAsync(context.OldXmlReader.Prefix, "color", context.OldXmlReader.NamespaceURI);
                    await context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, Diagonal.Color);
                    await context.NewXmlWriter.WriteEndElementAsync();
                }
                await context.NewXmlWriter.WriteEndElementAsync();
            }

            await context.NewXmlWriter.WriteEndElementAsync();
        }
    }
}
