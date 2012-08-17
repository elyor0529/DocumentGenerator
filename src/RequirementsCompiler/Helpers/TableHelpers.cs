// -----------------------------------------------------------------------
// <copyright file="TableHelpers.cs" company="">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace RequirementsCompiler.Helpers
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class TableHelpers
    {
        // Take the data from a two-dimensional array and build a table at the 
        // end of the supplied document.
        public static void AddTable(WordprocessingDocument document, string[,] data)
        {

            var doc = document.MainDocumentPart.Document;

            var table = new Table();

            var props = new TableProperties(
                new TableWidth()
                    {
                        Width = "100%"
                    },
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new BottomBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new LeftBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 0
                },
                new RightBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 0
                },
                new InsideHorizontalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                },
                new InsideVerticalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 1
                }));

            table.AppendChild(props);

            var numberOfRows = data.GetUpperBound(0);
            var numberOfColumns = data.GetUpperBound(1);

            var columnPercent = (100 / (numberOfColumns + 1)).ToString();

            for (var i = 0; i <= numberOfRows; i++)
            {
                var tr = new TableRow();
                for (var j = 0; j <= numberOfColumns; j++)
                {
                    var tc = new TableCell(new TableCellWidth() { Width = columnPercent });
                    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Pct }));

                    tr.Append(tc);
                }

                table.Append(tr);
            }

            doc.Body.Append(table);
        }
    }
}
