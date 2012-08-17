// -----------------------------------------------------------------------
// <copyright file="ParagraphExtensions.cs" company="">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace RequirementsCompiler.Helpers
{
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public static class ParagraphExtensions
    {
        public static Paragraph ApplyStyle(this Paragraph paragrah, string style)
        {
            ParagraphProperties properties = new ParagraphProperties();
            properties.ParagraphStyleId = new ParagraphStyleId() { Val = style };
            paragrah.Append(properties);
            return paragrah;
        }

        public static Paragraph AppendText(this Paragraph paragrah, string text)
        {
            paragrah.AppendChild(new Run()).AppendChild(new Text(text));
            return paragrah;
        }

        public static Paragraph AsBulletedList(this Paragraph paragraph, int level, int id)
        {
            NumberingProperties properties = new NumberingProperties();
            properties.NumberingLevelReference = new NumberingLevelReference() { Val = 0 };
            properties.NumberingId = new NumberingId() { Val = 1 };
            paragraph.Append(properties);
            return paragraph;
        }
    }
}
