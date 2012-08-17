// -----------------------------------------------------------------------
// <copyright file="StyleHelpers.cs" company="">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace RequirementsCompiler.Helpers
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using RequirementsCompiler.Maps;

    using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
    using Styles = DocumentFormat.OpenXml.Wordprocessing.Styles;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public static class StyleHelpers
    {       
        public static StyleDefinitionsPart ApplyStyle(this StyleDefinitionsPart part, IStyle mystyle)
        {
            var pRp = new RunProperties();
            var color = new Color()
            {
                Val = mystyle.Color
            };
            var fonts = new RunFonts
            {
                Ascii = mystyle.FontName
            };
            pRp.Append(color);
            pRp.Append(fonts);

            if (mystyle.Bold)
            {
                pRp.Append(new Bold());
            }
            
            pRp.Append(new FontSize()
            {
                Val = mystyle.FontSize.ToString()
            });

            var style = new Style { StyleId = mystyle.Id };
            style.Append(new Name() { Val = mystyle.Name });
            style.Append(new BasedOn() { Val = mystyle.BasedOn });
            style.Append(new NextParagraphStyle() { Val = "Normal" });
            
            style.Append(pRp);

            if (part.Styles == null)
            {
                part.Styles = new Styles();
            }

            part.Styles.Append(style);
            part.Styles.Save();

            return part;
        }

        // Create a new style with the specified styleid and stylename and add it to the specified
        // style definitions part.
        private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename)
        {
            // Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            // Create a new paragraph style and specify some of the properties.
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true
            };
            StyleName styleName1 = new StyleName() { Val = stylename };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts() { Ascii = "Lucida Console" };
            Italic italic1 = new Italic();
            // Specify a 12 point size.
            FontSize fontSize1 = new FontSize() { Val = "24" };
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }        
    }
}
