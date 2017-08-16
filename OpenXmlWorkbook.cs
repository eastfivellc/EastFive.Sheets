using BlackBarLabs.Extensions;
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

namespace EastFive.Sheets
{
    public class OpenXmlWorkbook : IUnderstandSheets
    {
        private SpreadsheetDocument workbook;

        public OpenXmlWorkbook(Stream stream)
        {
            workbook = SpreadsheetDocument.Open(stream, false);
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
        }

        public static IUnderstandSheets Create(System.IO.Stream stream)
        {
            using (var workbook = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                #region Custom properties


                #endregion

                //var worksheetsPart = workbook.AddWorkbookPart();
                var workbookPart = workbook.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var writer = OpenXmlWriter.Create(worksheetPart);
                var worksheet = new Worksheet();
                writer.WriteStartElement(worksheet);
                writer.WriteStartElement(new SheetData());

                #region Header Row

                var headerRow = new Row();
                var attributesHeaderRow = new OpenXmlAttribute[]
                {
                    // this is the row index
                    new OpenXmlAttribute("r", null, "1")
                };
                writer.WriteStartElement(headerRow, attributesHeaderRow);
                writer.WriteCell("ProductID (DO NOT MODIFY) Leave blank for new products");
                var headerCellsCategories = categoryTitles.Select(
                            (categoryTitle) => writer.WriteCell($"C:{categoryTitle}"))
                            .ToArray();
                writer.WriteCell("Product Name");
                var headerCellsProperties = propertyNames
                    .Select(
                        (property) => writer.WriteCell($"P({property.Key.ToString("N")}){property.Value}"))
                    .ToArray();
                // this is for header row
                writer.WriteEndElement();

                #endregion

                #region Body

                var rows = products.Select(
                    (result, index) =>
                    {
                        var attributeListRow = new OpenXmlAttribute[]
                        {
                            // this is the row index
                            new OpenXmlAttribute("r", null, (index+2).ToString("N"))
                        };

                        var row = new Row();
                        writer.WriteStartElement(row, attributeListRow);

                        writer.WriteCell(result.ProductId);
                        var categoryCells = result.Categories.Select(
                            (category) => writer.WriteCell($"C({category.ToString("N")}){categoryLookup[category]}"))
                            .ToArray();
                        writer.WriteCell(result.Name);
                        var cells = result.Properties.Select(
                            (property) => writer.WriteCell(result.Properties[property.Key])).ToArray();

                        // this is for Row
                        writer.WriteEndElement();
                        return row;
                    }).ToArray();

                #endregion

                // this is for SheetData
                writer.WriteEndElement();
                // this is for Worksheet
                writer.WriteEndElement();
                writer.Close();

                writer = OpenXmlWriter.Create(workbook.WorkbookPart);
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());

                // you can use object initialisers like this only when the properties
                // are actual properties. SDK classes sometimes have property-like properties
                // but are actually classes. For example, the Cell class has the CellValue
                // "property" but is actually a child class internally.
                // If the properties correspond to actual XML attributes, then you're fine.
                writer.WriteElement(new Sheet()
                {
                    Name = "Sheet1", // products.ProductTypeId.ToString("N"),
                    SheetId = 1,
                    Id = workbook.WorkbookPart.GetIdOfPart(worksheetPart)
                });

                writer.WriteEndElement(); // Write end for WorkSheet Element


                writer.WriteEndElement(); // Write end for WorkBook Element
                writer.Close();

                workbook.Close();

                var response = request.CreateResponse(HttpStatusCode.OK);
                stream.Position = 0;
                response.Content = new StreamContent(stream);
                response.Content.Headers.ContentType =
                    new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.template");
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = manufacturerId.HasValue ? $"{manufacturerId}.xlsx" : "products.xlsx",
                };
                return response;
            }
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
