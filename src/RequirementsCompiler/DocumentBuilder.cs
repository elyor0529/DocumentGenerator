// -----------------------------------------------------------------------
// <copyright file="DocumentBuilder.cs" company="">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace RequirementsCompiler
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Serialization;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using RequirementsCompiler.Helpers;
    using RequirementsCompiler.Maps;

    using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
    using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
    using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
    using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class DocumentBuilder
    {
        private readonly Modules _modules;

        public DocumentBuilder(string path)
        {
            var serializer = new XmlSerializer(typeof(Modules));
            var reader = new StreamReader(path);
            _modules = (Modules)serializer.Deserialize(reader);
        }

        public void Build()
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create("sample.docx", WordprocessingDocumentType.Document))
            {
                wordDocument.Initialize();

                MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                Body body = mainPart.Document.Body;

                var summaries = new Dictionary<Module, List<Dictionary<Tuple<Expertise, Priority>, decimal>>>();

                for (int i = 0; i < this._modules.Module.Length; i++)
                {
                    var module = this._modules.Module[i];
                    summaries.Add(module, new List<Dictionary<Tuple<Expertise, Priority>, decimal>>());


                    FormatTitle(body, module);
                    TableHelpers.AddTable(
                        wordDocument,
                        new[,]
                            {
                                {
                                    "Status", "Version", "Priority", "License"
                                },
                                { module.Status.ToString(), module.Version, module.Priority.ToString(), module.License.ToString() }
                            });

                    body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading2.Id).AppendText("Technologies"));

                    TableHelpers.AddTable(wordDocument, this.BuildTechnologyList(module.Technologies));

                    body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading2.Id).AppendText("Overview"));
                    body.AppendChild(new Paragraph()).AppendChild(new Run()).AppendChild(new Text(module.Overview));
                    body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading2.Id).AppendText("Dependencies"));
                    body.AppendChild(new Paragraph().AppendText("This module has dependencies on"));

                    foreach (var dep in module.Dependencies)
                    {
                        var paragraph = new Paragraph().AsBulletedList(0, 1);
                        paragraph.AppendText(dep);
                        body.AppendChild(paragraph);
                    }

                    body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading2.Id).AppendText("Work Remaining"));
                    TableHelpers.AddTable(
                       wordDocument,
                       this.BuildFeaturesList(module.Features));

                    var summary = new Dictionary<Tuple<Expertise, Priority>, decimal>();

                    body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading2.Id).AppendText("Pricing Matrix"));
                    TableHelpers.AddTable(
                       wordDocument,
                       this.BuildPricingMatrix(module, ref summary));

                    summaries[module].Add(summary);

                    body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading2.Id).AppendText("Totals"));
                    TableHelpers.AddTable(
                       wordDocument,
                       CalculateSummaryValues(summary, module));

                    body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                }

                body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading2.Id).AppendText("Summary"));
                TableHelpers.AddTable(wordDocument, this.CalculateTotalSummaries(summaries));
            }
        }

        public string[,] CalculateTotalSummaries(Dictionary<Module, List<Dictionary<Tuple<Expertise, Priority>, decimal>>> summaries)
        {
            var ret = new string[summaries.Keys.Count + 2, 5];

            ret[0, 0] = "Module Name";
            ret[0, 1] = "Hourly Estimate";
            ret[0, 2] = "License Discount";
            ret[0, 3] = "License Monthly";
            ret[0, 4] = "Total";

            decimal hourlyTotal = 0;
            decimal grandTotal = 0;
            decimal netTotal = 0;
            decimal discounts = 0;

            var currentRow = 1;
            var currentColumn = 0;
            foreach (var module in summaries.Keys)
            {
                decimal total = 0;
                decimal discount = 0;
                foreach (var x in summaries[module])
                {
                    total += x.Sum(y => y.Value * this.GetRate(y.Key.Item1) * this.GetModifier(y.Key.Item2));
                }

                discount = total - (this.GetLicenseModifier(module.License) * total);

                ret[currentRow, currentColumn] = module.Name;
                ret[currentRow, 1] = total.ToString("C");
                ret[currentRow, 2] = discount.ToString("C");
                ret[currentRow, 3] = 0.ToString("C");
                ret[currentRow, 4] = (total - discount).ToString("C");

                hourlyTotal += total;
                discounts += discount;

                currentRow++;
            }
            
            ret[currentRow, 1] = hourlyTotal.ToString("C");
            ret[currentRow, 2] = discounts.ToString("C");
            ret[currentRow, 2] = 
            ret[currentRow, 3] = (hourlyTotal - discounts).ToString("C");

            return ret;
        }

        public string[,] CalculateSummaryValues(Dictionary<Tuple<Expertise, Priority>, decimal> summary, Module module)
        {
            var values = new string[2, 4];
            decimal totalCost = 0;
            foreach (KeyValuePair<Tuple<Expertise, Priority>, decimal> value in summary)
            {
                totalCost += value.Value * this.GetRate(value.Key.Item1) * this.GetModifier(value.Key.Item2);
            }

            values[0, 0] = "Hourly Labor";
            values[0, 1] = "License Discount";
            values[0, 2] = "License (Monthly)";
            values[0, 3] = "Total";

            values[1, 0] = totalCost.ToString("C");
            values[1, 1] = 0.ToString("C");
            values[1, 2] = 0.ToString("C");
            values[1, 3] = totalCost.ToString("C");


            return values;
        }

        public decimal GetModifier(Priority priority)
        {
            switch (priority)
            {
                case Priority.Urgent:
                    return 1;
                    break;
                case Priority.High:
                    return 0.75M;
                    break;
                default:
                    return 0.5M;
            }
        }

        public decimal GetRate(Expertise expertise)
        {
            switch (expertise)
            {
                case Expertise.Expert:
                    return 100;
                    break;
                case Expertise.Novice:
                    return 25;
                    break;
                default:
                    return 35;
            }
        }

        private string[,] BuildPricingMatrix(
            Module module, ref Dictionary<Tuple<Expertise, Priority>, decimal> hourSummary)
        {
            var values = new string[Enum.GetValues(typeof(Priority)).GetUpperBound(0) + 3, Enum.GetValues(typeof(Expertise)).GetUpperBound(0) + 2];

            values[0, 0] = string.Empty;

            int currentRow = 0;
            int currentColumn = 0;
            decimal sum;

            hourSummary = new Dictionary<Tuple<Expertise, Priority>, decimal>();

            foreach (Expertise expertise in Enum.GetValues(typeof(Expertise)))
            {
                foreach (Priority priorty in Enum.GetValues(typeof(Priority)))
                {
                    var tuple = new Tuple<Expertise, Priority>(expertise, priorty);

                    hourSummary.Add(tuple, 0);

                    if (currentColumn == 0)
                    {
                        values[currentRow + 1, 0] = priorty.ToString() + " (" + this.GetModifier(priorty) + ")";
                    }

                    if (currentRow == 0)
                    {
                        values[0, currentColumn + 1] = expertise.ToString() + " @" + this.GetRate(expertise).ToString("C");
                    }

                    var hours = module.GetHours(priorty, expertise);
                    var rate = this.GetRate(expertise) * this.GetModifier(priorty);

                    hourSummary[tuple] += hours;

                    if (hours > 0)
                    {
                        values[currentRow + 1, currentColumn + 1] = hours.ToString("F") + " @" + rate.ToString("C");
                    }

                    currentRow++;
                }

                values[currentRow + 1, 0] = "Totals";

                decimal hrs = 0;
                decimal money = 0;
                foreach (var x in hourSummary)
                {
                    if (x.Key.Item1 == expertise)
                    {
                        hrs += x.Value;
                        money += x.Value * this.GetRate(x.Key.Item1) * this.GetModifier(x.Key.Item2);
                    }
                }

                values[currentRow + 1, currentColumn + 1] = hrs > 0
                                                                ? hrs.ToString("F") + " (" + money.ToString("C")
                                                                  + ")"
                                                                : string.Empty;

                currentColumn++;
                currentRow = 0;
            }

            return values;
        }

        public decimal GetLicenseModifier(License license)
        {
            switch(license)
            {
                case License.Proprietary:
                    return 1;
                    break;
                    case License.Scaffeine:
                    return 0.75M;
                    break;
                    case License.OpenSource:
                    return 0.5M;
                default:
                    return 0.5M;
            }
        }

        private string[,] BuildFeaturesList(ModuleFeature[] features)
        {
            var values = new string[features.Length + 1, 4];

            values[0, 0] = "Area";
            values[0, 1] = "Name";
            values[0, 2] = "Priority";
            values[0, 3] = "Hours";

            for (var i = 0; i < features.Length; i++)
            {
                values[i + 1, 0] = features[i].Area.ToString();
                values[i + 1, 1] = features[i].Name;
                values[i + 1, 2] = features[i].Priority.ToString();
                values[i + 1, 3] = features[i].Remaining.ToString();
            }

            return values;
        }

        private string[,] BuildTechnologyList(ModuleTechnology[] technologies)
        {
            var values = new string[technologies.Length + 1, 3];

            values[0, 0] = "Technology Name";
            values[0, 1] = "Amount of work";
            values[0, 2] = "Skill Level Required";

            for (var i = 0; i < technologies.Length; i++)
            {
                values[i + 1, 0] = technologies[i].Name;
                values[i + 1, 1] = technologies[i].Workload.ToString();
                values[i + 1, 2] = technologies[i].Expertise.ToString();
            }

            return values;
        }

        private void FormatTitle(Body body, Module module)
        {
            var para = body.AppendChild(new Paragraph().ApplyStyle(StyleDefinitions.Heading1.Id));

            var run = para.AppendChild(new Run());

            run.AppendChild(new Text(module.Name));
        }
    }
}
