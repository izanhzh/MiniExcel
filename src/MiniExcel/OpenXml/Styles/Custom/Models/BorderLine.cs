namespace MiniExcelLibs.OpenXml.Styles.Custom.Models
{
    public class BorderLine
    {
        public BorderStyle Style { get; set; }

        public string Color { get; set; }

        public BorderLine(BorderStyle style = BorderStyle.None, string color = null)
        {
            Style = style;
            Color = color;
        }
    }
}
