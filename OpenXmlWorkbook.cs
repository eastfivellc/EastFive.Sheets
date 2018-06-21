using BlackBarLabs.Extensions;
using EastFive.Linq;
using BlackBarLabs.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EastFive.Extensions;

namespace EastFive.Sheets
{
    public class OpenXmlWorkbook : IUnderstandSheets
    {
        private SpreadsheetDocument workbook;
        private WorkbookPart workbookPart = default(WorkbookPart);
        private IDictionary<string, WorksheetPart> sheets = new Dictionary<string, WorksheetPart>();

        public OpenXmlWorkbook(Stream stream)
        {
            workbook = SpreadsheetDocument.Open(stream, false);
        }

        private OpenXmlWorkbook(SpreadsheetDocument workbook)
        {
            this.workbook = workbook;
        }

        public static TResult Create<TResult>(System.IO.Stream stream,
            Func<IUnderstandSheets, TResult> callback)
        {
            using (var workbook = new OpenXmlWorkbook(SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)))
            {
                var result = callback(workbook);

                #region Create sheets that link to the workbook pages

                var oxw = OpenXmlWriter.Create(workbook.workbook.WorkbookPart);
                oxw.WriteStartElement(new Workbook());
                oxw.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                // you can use object initialisers like this only when the properties
                // are actual properties. SDK classes sometimes have property-like properties
                // but are actually classes. For example, the Cell class has the CellValue
                // "property" but is actually a child class internally.
                // If the properties correspond to actual XML attributes, then you're fine.

                UInt32Value sheetCount = 1;
                foreach (var sheet in workbook.sheets)
                {
                    var sheetCreated = new Sheet()
                    {
                        Name = sheet.Key,
                        SheetId = sheetCount,
                        Id = workbook.workbook.WorkbookPart.GetIdOfPart(sheet.Value)
                    };
                    oxw.WriteElement(sheetCreated);
                    sheetCount++;
                }

                // this is for Sheets
                oxw.WriteEndElement();
                // this is for Workbook
                oxw.WriteEndElement();
                oxw.Close();

                #endregion

                return result;
            }
        }

        public static TResult Read<TResult>(Stream stream,
            Func<IUnderstandSheets, TResult> onRead)
        {
            using (var workbook = new OpenXmlWorkbook(stream))
            {
                return onRead(workbook);
            }
        }

        public TResult WriteCustomProperties<TResult>(
            Func<Action<string, string>, TResult> callback)
        {
            var customPropsPart = workbook.AddCustomFilePropertiesPart();

            var properties = new CustomDocumentProperty[] { };
            var result = callback(
                (string name, string value) =>
                {
                    var property = new CustomDocumentProperty();
                    property.FormatId = Guid.NewGuid().ToString("B");
                    property.Name = name;
                    property.VTBString = new VTBString(value);
                    property.PropertyId = properties.Length + 2;
                    properties = properties.Append(property).ToArray();
                });

            customPropsPart.Properties = new Properties(properties);

            var writer = OpenXmlWriter.Create(customPropsPart);
            writer.WriteStartElement(customPropsPart.Properties);
            writer.WriteEndElement();
            writer.Close();
            return result;
        }

        public IEnumerable<ISheet> ReadSheets()
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
                                return new OpenXmlSheet(workbook, workbookPart, worksheet, worksheetData);
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

        public void Dispose()
        {
            

            this.workbook.Close();
            this.workbook.Dispose();
        }

        public struct CellReference
        {
            public string value;
            public string formula;
        }

        public TResult WriteSheetByRow<TResult>(Func<Action<object[]>, TResult> writeRowCallback, string sheetName = "Sheet1")
        {
            if (this.workbookPart.IsDefaultOrNull())
            {
                this.workbook.AddWorkbookPart();
                this.workbookPart = this.workbook.WorkbookPart;
            }

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var writer = OpenXmlWriter.Create(worksheetPart);

            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            var rowIndex = 1;
            var result = writeRowCallback(
                (rowCells) =>
                {
                    var row = new Row();
                    var attributes = new OpenXmlAttribute[]
                    {
                        // this is the row index
                        new OpenXmlAttribute("r", null, rowIndex.ToString())
                    };
                    writer.WriteStartElement(row, attributes);

                    foreach (var rowCell in rowCells)
                        WriteCell(writer, rowCell, rowIndex);

                    writer.WriteEndElement();
                    rowIndex++;
                });

            // this is for SheetData
            writer.WriteEndElement();
            // this is for Worksheet
            writer.WriteEndElement();
            writer.Close();

            sheets.Add(sheetName, worksheetPart);

            return result;
        }

        private static Cell WriteCell(OpenXmlWriter writer, object value, int rowIndex)
        {
            var cell = new Cell();
            //cell.CellReference = "";
            var attributeListCell = new OpenXmlAttribute[] { };

            if (default(object) == value)
                cell.DataType = CellValues.String;
            else if (typeof(bool).IsInstanceOfType(value))
                cell.DataType = CellValues.Boolean;
            else if (typeof(DateTime).IsInstanceOfType(value))
                cell.DataType = CellValues.Date;
            else if (typeof(double).IsAssignableFrom(value.GetType()))
                cell.DataType = CellValues.Number;
            else
                cell.DataType = CellValues.String;
            
            writer.WriteStartElement(cell);
            
            if (default(object) == value)
            {
                var cellValue = new CellValue("");
                writer.WriteElement(cellValue);
            }
            else if (typeof(CellReference).IsAssignableFrom(value.GetType()))
            {
                var cellRef = (CellReference)value;
                
                var formula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula(cellRef.formula);
                writer.WriteElement(formula);
                var cellValue = new CellValue(cellRef.value);
                writer.WriteElement(cellValue);
                // <c r="H2" t="str">
                //  <f>Manufacturer!A15</f>
                //  <v>30</v>
                //</c>
            } else
            {
                var cellValue = new CellValue(value.ToString());
                writer.WriteElement(cellValue);
            }

            writer.WriteEndElement();

            return cell;
        }
    }
}
