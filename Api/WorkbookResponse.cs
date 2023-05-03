using System;
using EastFive.Api;
using EastFive.Api.Sheets;
using EastFive.Reflection;

using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using ClosedXML.Excel;
using EastFive.Extensions;
using System.Linq;

namespace EastFive.Sheets.Api
{
    public interface IWriteSheet
    {
        string GetSheetName(Type resourceType);
        IHeaderData[] GetHeaderData(Type resourceType, Type[] otherTypes);
        void WriteSheet<TResource>(IXLWorksheet worksheet,
            TResource[] resource1s, Type[] types, IHeaderData[] headerDatas);
    }

    public interface IHeaderData
    {
        int ColumnsIndex { get;  }

        int ColumnsUsed { get; }

        IWriteCell CellWriter { get; }

        MemberInfo MemberInfo { get; }
    }

    public interface IWriteCell
    {
        IHeaderData ComputeHeader(Type resourceType, int colIndex, MemberInfo member, Type[] otherTypes);
        IXLCell WriteHeader<TResource>(IXLWorksheet worksheet, IHeaderData headerDataObj, Type[] otherTypes, TResource[] resources);
        IXLCell WriteCell<TResource>(IXLWorksheet worksheet, IHeaderData headerData, int rowIndex, TResource resource);
    }

    public class SheetCellStandardAttribute : Attribute, IWriteCell
    {
        public string ColumnName { get; set; }

        private struct HeaderData : IHeaderData
        {
            public int ColumnsIndex { get; set; }

            public int ColumnsUsed => HasLink ? 2 : 1;

            public MemberInfo MemberInfo { get; set; }

            public bool HasLink => this.externalSheetName.HasBlackSpace();

            public IWriteCell CellWriter { get; set; }

            public string externalSheetName;
        }

        public IXLCell WriteCell<TResource>(IXLWorksheet worksheet, IHeaderData headerDataObj, int rowIndex, TResource resource)
        {
            var headerData = (HeaderData)headerDataObj;
            var member = headerData.MemberInfo;
            var value = member.GetValue(resource);

            var memberType = member.GetPropertyOrFieldType();
            var cell = worksheet.Cell(rowIndex, headerData.ColumnsIndex);

            if (typeof(IReferenceableOptional).IsAssignableFrom(memberType))
            {
                var irefOptValue = (IReferenceableOptional)value;

                if (headerData.HasLink)
                    WriteRefLink(worksheet, cell, headerData, true);

                if (irefOptValue.HasValue)
                {
                    value = irefOptValue.id.Value.ToString();
                    var cellValueValue = XLCellValue.FromObject(value);
                    return cell.SetValue(cellValueValue);
                }

                var blankValue = Blank.Value;
                return cell.SetValue(blankValue);
            }

            if (typeof(IReferenceable).IsAssignableFrom(memberType))
            {
                var irefValue = (IReferenceable)value;
                value = irefValue.id.ToString();
                if(headerData.HasLink)
                {
                    WriteRefLink(worksheet, cell, headerData, false);
                }
            }

            var cellValue = XLCellValue.FromObject(value);
            var updatedCell = cell.SetValue(cellValue);
            return updatedCell;
        }

        private IXLCell WriteRefLink(IXLWorksheet worksheet, IXLCell sourceCell, HeaderData headerData, bool optional)
        {
            var sourceColumn = sourceCell.WorksheetColumn();
            var matchValueColNumber = sourceColumn.ColumnNumber();
            var matchValueColLetter = sourceColumn.ColumnLetter();
            var matchValueRow = sourceCell.WorksheetRow().RowNumber();

            var cellLink = worksheet.Cell(matchValueRow, matchValueColNumber + 1);
            var externalSheetName = headerData.externalSheetName; // "'Affirm Quality Measure Types'";
            var externalSheetPrefix = $"'{externalSheetName}'!";
            var externalColumnMatch = "A";
            var externalColumnDisplay = "B";
            var externalColumnLinkTo = externalColumnDisplay;
            var lookupRange = $"{externalSheetPrefix}{externalColumnMatch}:{externalColumnMatch}";
            
            var matchParam = $"(MATCH({matchValueColLetter}{matchValueRow},{lookupRange},0))";
            var linkedCellAddress = $"\"#\"&\"{externalSheetPrefix}{externalColumnLinkTo}\"&{matchParam}";
            var displayTextAddress = $"\"{externalSheetPrefix}{externalColumnDisplay}\"&{matchParam}";
            var linkDisplayText = $"INDIRECT({displayTextAddress})";
            var fullFormula = optional?
                $"=IFNA(HYPERLINK({linkedCellAddress},{linkDisplayText}))"
                :
                $"=HYPERLINK({linkedCellAddress},{linkDisplayText})";
            cellLink.SetFormulaA1(fullFormula);
            return cellLink;
        }

        public IXLCell WriteHeader<TResource>(IXLWorksheet worksheet, IHeaderData headerDataObj, Type[] otherTypes, TResource[] resources)
        {
            var headerData = (HeaderData)headerDataObj;
            var colIndex = headerData.ColumnsIndex;
            var member = headerData.MemberInfo;

            var value = GetValueToWrite();
            var cellValue = XLCellValue.FromObject(value);
            var cell = worksheet.Cell(1, colIndex);
            cell.SetValue(cellValue);

            if(headerData.HasLink)
            {
                var cellLink = worksheet.Cell(1, colIndex + 1);
                var cellLinkValue = XLCellValue.FromObject($"{headerData.externalSheetName} Link");
                cellLink.SetValue(cellLinkValue);
            }

            return cell;

            string GetValueToWrite()
            {
                if (ColumnName.HasBlackSpace())
                    return ColumnName;
                return member.Name;
            }
        }

        public IHeaderData ComputeHeader(Type resourceType, int colIndex, MemberInfo member, Type[] otherTypes)
        {
            var hd = new HeaderData
            {
                ColumnsIndex = colIndex,
                MemberInfo = member,
                CellWriter = this,
            };

            var memberType = member.GetPropertyOrFieldType();
            if (typeof(IReferenceable).IsAssignableFrom(memberType)
                || typeof(IReferenceableOptional).IsAssignableFrom(memberType))
            {
                var genericArgs = memberType.GetGenericArguments();
                if (!genericArgs.Any())
                    return hd;
                var refdType = genericArgs.First();
                if (refdType == resourceType)
                    return hd;

                if (!otherTypes.Contains(refdType))
                    return hd;

                if (!refdType.TryGetAttributeInterface(out IWriteSheet externalReferenceWriter))
                    return hd;

                hd.externalSheetName = externalReferenceWriter.GetSheetName(refdType);
            }

            return hd;
        }
    }

    public class WriteSheetSerializedAttribute : Attribute, IWriteSheet
    {
        public string SheetName { get; set; }

        public IHeaderData[] GetHeaderData(Type resourceType, Type[] otherTypes)
        {
            return resourceType
                .GetPropertyAndFieldsWithAttributesInterface<IWriteCell>()
                .Aggregate(
                    new IHeaderData[] { },
                    (headerDatas, tpl) =>
                    {
                        var index = headerDatas.Select(hd => hd.ColumnsUsed).Sum();
                        var (member, cellWriter) = tpl;
                        var headerData = cellWriter.ComputeHeader(resourceType, index + 1, member, otherTypes);
                        return headerDatas.Append(headerData).ToArray();
                    })
                .ToArray();
        }

        public string GetSheetName(Type resourceType)
        {
            return SheetName.HasBlackSpace()?
                SheetName
                :
                resourceType.FullName;
        }

        public void WriteSheet<TResource>(IXLWorksheet worksheet,
            TResource[] resources, Type[] otherTypes, IHeaderData[] headerDatas)
        {
            var headerCells = headerDatas
                .Select(
                    hd => hd.CellWriter.WriteHeader(worksheet, hd, otherTypes, resources))
                .ToArray();

            IXLCell[] allCells = resources
                .SelectMany(
                    (resource, index) =>
                    {
                        return headerDatas
                            .Select(
                                headerData =>
                                {
                                    var cellWriter = headerData.CellWriter;
                                    var cell = cellWriter.WriteCell<TResource>(worksheet, headerData, index + 2, resource);
                                    return cell;
                                });

                    })
                .Concat(headerCells)
                .ToArray();
        }
    }

    [WorkbookResponse]
    public delegate IHttpResponse WorkbookResponse<TResource1, TResource2, TResource3, TResource4>(
            TResource1[] resource1s, TResource2[] resource2s, TResource3[] resource3s, TResource4[] resource4s,
            string filename = "");

    public class WorkbookResponseAttribute : HttpGenericDelegateAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        [InstigateMethod]
        public IHttpResponse ContentResponse<TResource1, TResource2, TResource3, TResource4>(
            TResource1[] resource1s, TResource2[] resource2s, TResource3[] resource3s, TResource4[] resource4s,
            string filename = "")
        {
            var httpApiApp = this.httpApp as IApiApplication;
            var response = new WorkbookResponse<TResource1, TResource2, TResource3, TResource4>(
                resource1s:resource1s, resource2s:resource2s, resource3s:resource3s, resource4s:resource4s,
                filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class WorkbookResponse<TResource1, TResource2, TResource3, TResource4> : EastFive.Api.HttpResponse
        {
            TResource1[] resource1s; TResource2[] resource2s;
            TResource3[] resource3s; TResource4[] resource4s;

            public WorkbookResponse(
                    TResource1[] resource1s, TResource2[] resource2s,
                    TResource3[] resource3s, TResource4[] resource4s,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(request, HttpStatusCode.OK)
            {
                this.resource1s = resource1s;
                this.resource2s = resource2s;
                this.resource3s = resource3s;
                this.resource4s = resource4s;

                this.SetFileHeaders(
                    fileName.IsNullOrWhiteSpace() ? $"sheet.xlsx" : fileName,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                    true);
            }

            public override async Task WriteResponseAsync(Stream responseStream)
            {
                try
                {
                    var xlsData = ConvertToXlsxStreamAsync(
                        resource1s: resource1s, resource2s: resource2s,
                        resource3s: resource3s, resource4s: resource4s);

                    await responseStream.WriteAsync(xlsData, 0, xlsData.Length,
                        this.Request.CancellationToken);
                }
                catch (Exception ex)
                {

                }
            }

            private byte[] ConvertToXlsxStreamAsync(
                    TResource1[] resource1s, TResource2[] resource2s,
                    TResource3[] resource3s, TResource4[] resource4s)
            {
                using (var stream = new MemoryStream())
                {
                    var wb = new XLWorkbook();
                    WriteSheet<TResource1>(resource1s, new Type[] { typeof(TResource2), typeof(TResource3), typeof(TResource4) });
                    WriteSheet<TResource2>(resource2s, new Type[] { typeof(TResource1), typeof(TResource3), typeof(TResource4) });
                    WriteSheet<TResource3>(resource3s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource4) });
                    WriteSheet<TResource4>(resource4s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource3) });
                    wb.SaveAs(stream);

                    var buffer = stream.ToArray();
                    return buffer;

                    IWriteSheet GetSheetWriter(Type type)
                    {
                        if (!type.TryGetAttributeInterface(out IWriteSheet sheetWriter))
                            throw new Exception($"Cannot return objects of type {type.FullName} from {this.GetType().FullName}"
                                + $" without attribute implementing {nameof(IWriteSheet)}.");
                        return sheetWriter;
                    }

                    void WriteSheet<TResource>(TResource[] resources, Type[] otherTypes)
                    {
                        if (resources.IsDefaultOrNull())
                            return;

                        var sheetWriter = GetSheetWriter(typeof(TResource));
                        var sheetName = sheetWriter.GetSheetName(typeof(TResource));
                        var ws = wb.Worksheets.Add(sheetName);
                        var headerData = sheetWriter.GetHeaderData(typeof(TResource), otherTypes);
                        sheetWriter.WriteSheet(ws, resources, otherTypes, headerData);

                    }
                }
            }

        }
    }
}

