using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

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

        private Cell BuildEmptyCell(string reference)
        {
            var cell = new Cell();
            cell.CellValue = new CellValue("");
            cell.CellReference = new StringValue(reference);
            return cell;
        }

        public IEnumerable<string[]> ReadRows(Func<Type, object,
            Func<string>, string> serializer = default,
            bool discardAutoDecodeEncoding = default,
            Encoding[] discardEncodingsToUse = default)
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

                                    if(!TryGetDateFormat(out var dateFormatString))
                                        return (cellFormat, dateFormatString, false, false);

                                    var (shouldUse, isDateFormat) = IsUsable();
                                    return (cellFormat, dateFormatString, shouldUse, isDateFormat);

                                    (bool, bool) IsUsable()
                                    {
                                        string testOutput;
                                        try
                                        {
                                            testOutput = testTime.ToString(dateFormatString);
                                        }
                                        catch (FormatException)
                                        {
                                            return (false, false);
                                        }
                                        if (!DateTime.TryParseExact(testOutput, dateFormatString,
                                                System.Globalization.CultureInfo.InvariantCulture,
                                                System.Globalization.DateTimeStyles.None, out DateTime crossCheck))
                                        {
                                            return (false, true);
                                        }

                                        var isUsable = testTime.EqualToDay(crossCheck);
                                        return (isUsable, true);
                                    }

                                    bool TryGetDateFormat(out string dateFormatString)
                                    {
                                        dateFormatString = default;
                                        if (NumberFormatDictionary.TryGetValue(cellFormat.NumberFormatId, out string numberFormatString))
                                            return false;

                                        if (stylesLookup.TryGetValue(cellFormat.NumberFormatId, out var formatString))
                                        {
                                            if (OpenXmlToDotNetStyleLookup.TryGetValue(formatString.Value, out string dotNetFormatString))
                                            {
                                                dateFormatString = dotNetFormatString;
                                                return true;
                                            }

                                            // check for # or 0 found in number formats
                                            if (formatString.Value.IndexOfAny(new[] { '#', '0'}) != -1)
                                                return false;

                                            dateFormatString = formatString.Value;
                                            return true;
                                        }

                                        return DateFormatDictionary.TryGetValue(cellFormat.NumberFormatId, out dateFormatString);
                                    }
                                })
                            .ToArray();
                    },
                    () =>
                    {
                        return new (CellFormat, string, bool, bool) [] { };
                    });

            var rowsFromWorksheet = worksheetData
                .Descendants<Row>()
                .ToArray();

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
                        if (cell.IsDefaultOrNull() ||
                            (cell.CellValue.IsDefaultOrNull() && cell.InnerText.IsDefaultOrNull()))
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
                                        var (cf, styleText, shouldUse, isDateFormat) = styles[index];
                                        if (isDateFormat)
                                        {
                                            if (IsStyled())
                                            {
                                                var cellValueText = GetCellValue();
                                                if (double.TryParse(cellValueText, out var oaDate))
                                                {
                                                    try
                                                    {
                                                        var date = DateTime.FromOADate(oaDate);
                                                        if (serializer.IsNotDefaultOrNull())
                                                            return serializer(typeof(DateTime), date,
                                                                () =>
                                                                {
                                                                    return GetDTValue();
                                                                });

                                                        return GetDTValue();
                                                        string GetDTValue()
                                                        {
                                                            if (shouldUse)
                                                                return date.ToString(styleText);

                                                            if (date.Hour == 0)
                                                                if (date.Minute == 0)
                                                                    if (date.Second == 0)
                                                                        return date.ToShortDateString();
                                                            return date.ToString("yyyy/MM/dd HH:mm:ss");
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        ex.GetType();
                                                        return cellValueText;
                                                    }
                                                }
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

                            return GetCellValue();

                            string GetCellValue()
                            {
                                if (cell.DataType.IsNotDefaultOrNull())
                                {
                                    if (cell.DataType.HasValue)
                                    {
                                        var dataType = cell.DataType.InnerText;
                                        if (String.Equals(dataType, "s", StringComparison.OrdinalIgnoreCase)) // Is int
                                        {
                                            int sharedStringIndex;
                                            if (int.TryParse(cell.CellValue.Text, out sharedStringIndex))
                                                if (sharedStringIndex < sharedStrings.Length)
                                                    return sharedStrings[sharedStringIndex];
                                        }
                                        else if (String.Equals(dataType, "str", StringComparison.OrdinalIgnoreCase)) // Is string
                                        {
                                            return cell.CellValue.Text;
                                        }
                                        else if (String.Equals(dataType, "inlineStr", StringComparison.OrdinalIgnoreCase))
                                        {
                                            return cell.InnerText;
                                        }
                                    }
                                }
                                if (cell.CellValue.IsDefaultOrNull())
                                    return string.Empty;
                                return cell.CellValue.Text;
                            }
                        }
                        catch (Exception ex)
                        {
                            var type = ex.GetType();
                            return cell.CellValue.Text;
                        }
                    }
                })
                .ToArray();
            return rows;
        }

        // https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)
        private static readonly Dictionary<uint, string> DateFormatDictionary = new Dictionary<uint, string>()
        {
            [14] = "dd/MM/yyyy",
            [15] = "d-MMM-yy",
            [16] = "d-MMM",
            [17] = "MMM-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "M/d/yy h:mm",
            [30] = "M/d/yy",
            [34] = "yyyy-MM-dd",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [51] = "MM-dd",
            [52] = "yyyy-MM-dd",
            [53] = "yyyy-MM-dd",
            [55] = "yyyy-MM-dd",
            [56] = "yyyy-MM-dd",
            [58] = "MM-dd",
            [165] = "M/d/yy",
            [166] = "dd MMMM yyyy",
            [167] = "dd/MM/yyyy",
            [168] = "dd/MM/yy",
            [169] = "d.M.yy",
            [170] = "yyyy-MM-dd",
            [171] = "dd MMMM yyyy",
            [172] = "d MMMM yyyy",
            [173] = "M/d",
            [174] = "M/d/yy",
            [175] = "MM/dd/yy",
            [176] = "d-MMM",
            [177] = "d-MMM-yy",
            [178] = "dd-MMM-yy",
            [179] = "MMM-yy",
            [180] = "MMMM-yy",
            [181] = "MMMM d, yyyy",
            [182] = "M/d/yy hh:mm t",
            [183] = "M/d/y HH:mm",
            [184] = "MMM",
            [185] = "MMM-dd",
            [186] = "M/d/yyyy",
            [187] = "d-MMM-yyyy"
        };

        private static readonly Dictionary<uint, string> NumberFormatDictionary = new Dictionary<uint, string>()
        {
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [37] = "#,##0 ;(#,##0)",
            [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)",
            [40] = "#,##0.00;[Red](#,##0.00)",
            [48] = "##0.0E+0",
        };

        private static readonly Dictionary<string, string> OpenXmlToDotNetStyleLookup = new Dictionary<string, string>
        {
            { "yyyy/mm/dd", "yyyy/MM/dd" },
            { "[$-409]mmm\\ dd\\,\\ yyyy;@", "MM/dd/yyyy" },
            { "[$-0409]MMM dd, yyyy;@", "MM/dd/yyyy" },
            { "[$-10409]m/d/yyyy", "MM/dd/yyyy" },
            { "[$-10409]mm/dd/yyyy", "MM/dd/yyyy" },
            { "[$-10409]h:mm:ss\\ AM/PM",  "h:mm:ss tt"},
        };

        public void WriteRows(string fileName, object[] rows)
        {
            throw new NotImplementedException();
        }
    }
}
