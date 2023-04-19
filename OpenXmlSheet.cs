using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

using EastFive.Collections.Generic;
using EastFive.Linq;
using EastFive.Extensions;

namespace EastFive.Sheets
{
    class OpenXmlSheet : ISheet
    {
        private SpreadsheetDocument workbook;
        private WorkbookPart workbookPart;
        private Worksheet worksheetData;
        private Sheet worksheet;

        public string Name => worksheet.Name;

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
            var sharedStringsParts = workbookPart.GetPartsOfType<SharedStringTablePart>().ToArray();
            var sharedStrings = (sharedStringsParts.Count() > 0) ?
                sharedStringsParts.First().SharedStringTable
                    .Elements<SharedStringItem>()
                    .Select(
                        (item) =>
                        {
                            return item.InnerText;
                        })
                    .ToArray()
                :
                new string[] { };

            
            var styles = workbookPart
                .GetPartsOfType<WorkbookStylesPart>()
                .First(
                    (sytlePart, next) =>
                    {
                        var stylesLookup = sytlePart.Stylesheet.NumberingFormats
                            .NullToEmpty()
                            .Select(
                                (nf, index) =>
                                {
                                    var numberingFormat = (NumberingFormat)nf;
                                    return numberingFormat.NumberFormatId.PairWithValue(numberingFormat.FormatCode);
                                })
                            .ToDictionary();

                        var testTime = new DateTime(1, 2, 3, 4, 5, 6);

                        return sytlePart.Stylesheet.CellFormats
                            .Select(
                                cf =>
                                {
                                    var cellFormat = (CellFormat)cf;
                                    StringValue numberFormat;
                                    if (!stylesLookup.TryGetValue(cellFormat.NumberFormatId, out numberFormat))
                                        return (cellFormat, default(string), false);

                                    return (cellFormat, numberFormat.Value, IsUsable());

                                    bool IsUsable()
                                    {
                                        var testOutput = testTime.ToString(numberFormat);
                                        if (!DateTime.TryParseExact(testOutput, numberFormat,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out DateTime crossCheck))
                                        {
                                            return false;
                                        }

                                        var isUsable = testTime == crossCheck;
                                        return isUsable;
                                    }

                                })
                            .ToArray();
                    },
                    () =>
                    {
                        return new (CellFormat, string, bool) [] { };
                    });

            var rowsFromWorksheet = worksheetData
                .Descendants<Row>()
                .ToArray();

            //var allColumnIndexes = rowsFromWorksheet
            //    .SelectMany(
            //        row =>
            //        {
            //            if(!row.Spans.IsDefaultOrNull())
            //                return row.Spans.Items.SelectMany(rowSpan => Extract(rowSpan)).Append(row.Elements<Cell>().Count());
            //            if (!row.ChildElements.IsDefaultOrNull())
            //                return row.ChildElements.Count().AsEnumerable().Append(1);
            //            return new int[] { };
            //        })
            //    .ToArray();
            //var startIndex = allColumnIndexes.Min();
            //var lastIndex = allColumnIndexes.Max();
            //var columnIndexes = Enumerable.Range(startIndex, (lastIndex - startIndex) + 1);

            var allColumnNames = rowsFromWorksheet
                .SelectMany(
                    row =>
                    {
                        var cellOptions = row.Elements<Cell>()
                            .Select(cell => new String(cell.CellReference.Value.Where(c => c >= 'A' && c <= 'z').ToArray()))
                            .ToArray();
                        return cellOptions;
                    })
                .Distinct()
                .OrderBy(
                    str =>
                    {
                        var total = str
                            .Reverse()
                            .Select((c, i) => (c, (int)Math.Pow(100, i)))
                            .Aggregate(0,
                                (penalty, tpl) =>
                                {
                                    var (c, index) = tpl;
                                    var cOffset = c - ('A' - 1);
                                    return penalty + (index * cOffset);
                                });
                        return total;
                    })
                .ToArray();

            var rows = rowsFromWorksheet
                .Select(row =>
                {
                    //var columns = columnIndexes
                    //    .Select(colIndex => $"{(char)('A' + (colIndex - 1))}{row.RowIndex.ToString()}")
                    //    .ToArray();
                    var columns = allColumnNames
                        .Select(cn => $"{cn}{row.RowIndex}")
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
                        .Select(ConvertCell)
                        .ToArray();

                    string ConvertCell(Cell cell)
                    {
                        if (cell.IsDefaultOrNull() || cell.CellValue.IsDefaultOrNull())
                            return string.Empty;

                        try
                        {
                            if (cell.StyleIndex.IsNotDefaultOrNull())
                            {
                                if (cell.StyleIndex.HasValue)
                                {
                                    var index = cell.StyleIndex.Value;
                                    if (index < styles.Length)
                                    {
                                        var (cf, styleText, shouldUse) = styles[index];
                                        if (IsStyled())
                                        {
                                            if (double.TryParse(cell.CellValue.Text, out var oaDate))
                                            {
                                                var date = DateTime.FromOADate(oaDate);
                                                if (shouldUse)
                                                    return date.ToString(styleText);

                                                if (date.Hour == 0)
                                                    if (date.Minute == 0)
                                                        if (date.Second == 0)
                                                            return date.ToShortDateString();
                                                return date.ToString("yyyy/MM/dd HH:mm:ss");
                                            }
                                        }

                                        bool IsStyled()
                                        {
                                            if (styleText.HasBlackSpace())
                                                return true;

                                            if (cf.ApplyNumberFormat.IsDefaultOrNull())
                                                return false;
                                            if (cf.ApplyNumberFormat.HasValue)
                                                if (cf.ApplyNumberFormat.Value)
                                                    return true;

                                            return false;
                                        }
                                    }
                                }
                            }

                            if(cell.DataType.IsNotDefaultOrNull())
                            {
                                if(cell.DataType.HasValue)
                                {
                                    var dataType = cell.DataType.InnerText;
                                    if (String.Equals(dataType, "s", StringComparison.OrdinalIgnoreCase)) // Is int
                                    {
                                        int sharedStringIndex;
                                        if (int.TryParse(cell.CellValue.Text, out sharedStringIndex))
                                            if (sharedStringIndex < sharedStrings.Length)
                                                return sharedStrings[sharedStringIndex];
                                    }
                                    if (String.Equals(dataType, "str", StringComparison.OrdinalIgnoreCase)) // Is string
                                    {
                                        return cell.CellValue.Text;
                                    }
                                }
                            }
                            return cell.CellValue.Text;
                        }
                        catch (Exception ex)
                        {
                            var type = ex.GetType();
                            return cell.CellValue.Text;
                        }

                        //foreach (var attribute in cell.GetAttributes())
                        //    {
                        //        if (string.Equals(attribute.LocalName, "t", StringComparison.OrdinalIgnoreCase))
                        //        {
                        //            var typeAttr = attribute;
                        //            if (typeAttr.Value == "s") // Is int
                        //            {
                        //                int sharedStringIndex;
                        //                if (int.TryParse(cell.CellValue.Text, out sharedStringIndex))
                        //                    if (sharedStringIndex < sharedStrings.Length)
                        //                        return sharedStrings[sharedStringIndex];
                        //            }
                        //            if (typeAttr.Value == "str")
                        //            {
                        //                return cell.CellValue.Text;
                        //            }
                        //        }
                        //        if(string.Equals(attribute.LocalName, "s", StringComparison.OrdinalIgnoreCase))
                        //        {
                        //            if (attribute.Value == "1") // Is date
                        //            {
                        //                if (double.TryParse(cell.CellValue.Text, out var oaDate))
                        //                {
                        //                    var date = DateTime.FromOADate(oaDate);
                        //                    if (date.Hour == 0)
                        //                        if (date.Minute == 0)
                        //                            if (date.Second == 0)
                        //                                return date.ToShortDateString();
                        //                    return date.ToString("yyyy/MM/dd HH:mm:ss");
                        //                }
                        //            }
                        //        }
                        //        continue;
                        //    }
                    }
                })
                .ToArray();
            return rows;
        }

        public void WriteRows(string fileName, object[] rows)
        {
            throw new NotImplementedException();
        }
    }
}
