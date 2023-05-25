using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

                                    if(!TryGetDateFormat(out var dateFormatString))
                                        return (cellFormat, dateFormatString, false, false);

                                    var (canUseFormatting, isNumber) = IsUsable();
                                    return (cellFormat, dateFormatString, canUseFormatting, isNumber);

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
                                        if (stylesLookup.TryGetValue(cellFormat.NumberFormatId, out var numberFormat))
                                        {
                                            dateFormatString = numberFormat.Value;
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
                                        var (cf, styleText, shouldUse, isDateFormat) = styles[index];
                                        if (isDateFormat)
                                        {
                                            if (IsStyled())
                                            {
                                                if (double.TryParse(cell.CellValue.Text, out var oaDate))
                                                {
                                                    try
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
                                                    catch (Exception ex)
                                                    {
                                                        ex.GetType();
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
                    }
                })
                .ToArray();
            return rows;
        }

        // https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)
        private readonly Dictionary<uint, string> DateFormatDictionary = new Dictionary<uint, string>()
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

        public void WriteRows(string fileName, object[] rows)
        {
            throw new NotImplementedException();
        }
    }
}
