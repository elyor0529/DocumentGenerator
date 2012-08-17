// -----------------------------------------------------------------------
// <copyright file="DocumentHelpers.cs" company="">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace RequirementsCompiler.Helpers
{
    using System.Collections.Generic;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using RequirementsCompiler.Maps;

    using Header = DocumentFormat.OpenXml.Wordprocessing.Header;
    using NumberingFormat = DocumentFormat.OpenXml.Wordprocessing.NumberingFormat;
    using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
    using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public static class DocumentHelpers
    {
        public static WordprocessingDocument Initialize(this WordprocessingDocument document)
        {
            MainDocumentPart mainDocumentPart = document.AddMainDocumentPart();

            mainDocumentPart.Document = new Document();
            mainDocumentPart.Document.AppendChild(new Body());
            mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
            mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

            var headerPart = mainDocumentPart.AddNewPart<HeaderPart>();
            var footerPart = mainDocumentPart.AddNewPart<FooterPart>();
            var stylesPart = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            var numbersPart = mainDocumentPart.AddNewPart<NumberingDefinitionsPart>("asdf");
        
            string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
            string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);

            GenerateHeaderPartContent(headerPart);
            GenerateFooterPartContent(footerPart);
            GenerateNumbersPartContent(numbersPart);

            stylesPart.ApplyStyle(StyleDefinitions.Heading1);
            stylesPart.ApplyStyle(StyleDefinitions.Heading2);

            // Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
            IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

            foreach (var section in sections)
            {
                // Delete existing references to headers and footers
                section.RemoveAllChildren<HeaderReference>();
                section.RemoveAllChildren<FooterReference>();

                // Create the new header and footer reference node
                section.PrependChild(new HeaderReference() { Id = headerPartId });
                section.PrependChild(new FooterReference() { Id = footerPartId });
            }

            return document;
        }

        private static void GenerateNumbersPartContent(NumberingDefinitionsPart numbersPart)
        {
            var element =
                new Numbering(
                    new AbstractNum(
                        new Level(
                            new NumberingFormat()
                                {
                                    Val = NumberFormatValues.Bullet
                                },
                            new LevelText()
                                {
                                    Val = "·"
                                })
                            {
                                LevelIndex = 0
                            })
                            { AbstractNumberId = 1 });

            element.Save(numbersPart);
        }

        private static void GenerateHeaderPartContent(HeaderPart part)
        {
            var header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };

            var paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            var paragraphProperties1 = new ParagraphProperties();
            var paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            var run1 = new Run();
            var text1 = new Text { Text = "Header" };

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            header1.Append(paragraph1);

            part.Header = header1;
        }

        private static void GenerateFooterPartContent(FooterPart part)
        {
            var footer1 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };

            var paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            var paragraphProperties1 = new ParagraphProperties();
            var paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties1.Append(paragraphStyleId1);

            var run1 = new Run();
            var text1 = new Text { Text = "Footer" };

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            footer1.Append(paragraph1);

            part.Footer = footer1;
        }
    }
}
