using BlackBarLabs.Extensions;
using BlackBarLabs.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EastFive.Sheets
{
    class OpenXmlWorkbook
    {
        private SpreadsheetDocument workbook;

        public OpenXmlWorkbook(Stream stream)
        {
            workbook = SpreadsheetDocument.Open(stream, false);
        }

        public IEnumerable<ISheet> GetSheets()
        {
            return workbook
                .GetPartsOfType<WorkbookPart>()
                .SelectMany(
                    (workbookPart) => workbookPart.Workbook
                        .Descendants<Sheet>()
                        .Select(
                            (worksheet) =>
                            {
                                var wsPart = (WorksheetPart)(workbookPart.GetPartById(worksheet.Id));
                                var worksheetData = wsPart.Worksheet;
                                return new OpenXmlSheet(worksheetData);
                            }))
                .Select(openXmlSheet => (ISheet)openXmlSheet);
        }

        public TResult ReadCustomValues<TResult>(Func<KeyValuePair<string, string>[], TResult> onResults)
        {
            var customFileErrors = workbook
                .GetPartsOfType<CustomFilePropertiesPart>()
                .Select(
                    (customPropsPart) =>
                    {
                        var properties = customPropsPart.Properties
                            .Where(prop => prop is CustomDocumentProperty)
                            .Select(prop => prop as CustomDocumentProperty)
                            .ToArray();
                        var propertiesKvp = properties
                            .Select(prop => prop.Name.InnerText.PairWithValue(prop.VTBString.Text))
                            .ToArray();
                        return propertiesKvp;
                    })
                .SelectMany()
                .ToArray();
            return onResults(customFileErrors);
        }
    }
}
