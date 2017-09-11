using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;
using BlackBarLabs.Extensions;
using BlackBarLabs.Collections.Generic;
using BlackBarLabs.Linq;

namespace EastFive.Sheets
{
    class OpenXmlSheet : ISheet
    {
        private SpreadsheetDocument workbook;
        private WorkbookPart workbookPart;
        private Worksheet worksheetData;
        private Sheet worksheet;

        public OpenXmlSheet(SpreadsheetDocument workbook, WorkbookPart workbookPart, Sheet worksheet, Worksheet worksheetData)
        {
            this.worksheetData = worksheetData;
            this.workbook = workbook;
            this.workbookPart = workbookPart;
            this.worksheet = worksheet;
        }

        private int[] Extract(StringValue sv)
        {
            var regex = new Regex("([0-9]+):([0-9]+)");
            var match = regex.Match(sv.Value);
            if (!match.Success)
                return new int[] { };
            if (match.Groups.Count != 3)
                return new int[] { };
            var start = int.Parse(match.Groups[1].Value);
            var end = int.Parse(match.Groups[2].Value);
            return Enumerable.Range(start, (end - start)+1).ToArray();
        }

        private Cell BuildEmptyCell(string reference)
        {
            var cell = new Cell();
            cell.CellValue = new CellValue("");
            cell.CellReference = new StringValue(reference);
            return cell;
        }

        public IEnumerable<string[]> ReadRows()
        {
            var sharedStringsParts = workbookPart.GetPartsOfType<SharedStringTablePart>();
            var sharedStrings = (sharedStringsParts.Count() > 0) ?
                sharedStringsParts.First().SharedStringTable
                    .Elements<SharedStringItem>()
                    .Select(
                        (item) => item.InnerText)
                    .ToArray()
                :
                new string[] { };

            var rowsFromWorksheet = worksheetData
                .Descendants<Row>()
                .ToArray();

            var allColumnIndexes = rowsFromWorksheet
                .SelectMany(
                    row => row.Spans.Items.SelectMany(rowSpan => Extract(rowSpan)).Append(row.Elements<Cell>().Count()))
                .ToArray();
            var startIndex = allColumnIndexes.Min();
            var lastIndex = allColumnIndexes.Max();
            var columnIndexes = Enumerable.Range(startIndex, (lastIndex - startIndex) + 1);

            var rows = rowsFromWorksheet
                .Select(row =>
                {
                    var columns = columnIndexes
                        .Select(colIndex => $"{(char)('A' + (colIndex - 1))}{row.RowIndex.ToString()}")
                        .ToArray();
                    var elements = row.Elements<Cell>().ToArray();
                    if (columns.Length > 0 && columns.Length != elements.Length)
                    {
                        var elementLookup = elements
                            .Select(element => element.CellReference.Value.PairWithValue(element))
                            .ToDictionary();
                        elements = columns
                            .Select(reference => elementLookup.ContainsKey(reference) ?
                                elementLookup[reference]
                                :
                                BuildEmptyCell(reference))
                            .ToArray();
                    }

                    return elements
                        .Select(
                            (cell) =>
                            {
                                if (cell.IsDefaultOrNull() || cell.CellValue.IsDefaultOrNull())
                                    return string.Empty;
                                
                                if (!cell.HasAttributes)
                                    return cell.CellValue.Text;

                                try
                                {
                                    foreach (var attribute in cell.GetAttributes())
                                    {
                                        if (string.Compare(attribute.LocalName, "t", true) != 0)
                                            continue;
                                        var typeAttr = attribute;
                                        if (typeAttr.Value == "s")
                                        {
                                            int sharedStringIndex;
                                            if (int.TryParse(cell.CellValue.Text, out sharedStringIndex))
                                                if (sharedStringIndex < sharedStrings.Length)
                                                    return sharedStrings[sharedStringIndex];
                                        }
                                        continue;
                                    }
                                    return cell.CellValue.Text;
                                }
                                catch (Exception ex)
                                {
                                    return cell.CellValue.Text;
                                }
                            })
                        .ToArray();
                })
                .ToArray();
            return rows;
        }
    }
}
