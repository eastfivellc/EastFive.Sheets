using System;
using ClosedXML.Excel;
using EastFive.Sheets.Api;
using System.Linq;
using System.Collections.Generic;
using Irony;
using EastFive.Linq;

namespace EastFive.Sheets
{
    public class SerializedSheetAttribute : Attribute, IWriteSheet, IReadSheet
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
            return SheetName.HasBlackSpace() ?
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

        public IEnumerable<TResource> ReadSheet<TResource>(IXLWorksheet worksheet, IEnumerable<IXLRow> rows)
        {
            var remainingRows = rows.Pop(out IXLRow headerRow);

            var parsers = typeof(TResource)
                .GetPropertyAndFieldsWithAttributesInterface<IReadCell>()
                .Select(
                    (tpl) =>
                    {
                        var (member, cellReader) = tpl;
                        var headerMatchData = cellReader.MatchHeaders<TResource>(headerRow, member);
                        return (headerMatchData, cellReader);
                    })
                .ToArray();

            return remainingRows
                .Where(row => row.Cells().Any(cell => !cell.Value.IsBlank))
                .Select(
                    row =>
                    {
                        return parsers
                            .Aggregate(
                                Activator.CreateInstance<TResource>(),
                                (resource, tpl) =>
                                {
                                    var (headerMatchData, cellReader) = tpl;
                                    return cellReader.ReadCell(resource, row, headerMatchData);
                                });
                    });
        }
    }
}

