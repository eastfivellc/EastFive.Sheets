using System;
using ClosedXML.Excel;
using EastFive.Sheets.Api;
using System.Reflection;
using EastFive.Reflection;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using EastFive.Linq;

namespace EastFive.Sheets
{
    public class SheetCellStandardAttribute : Attribute, IWriteCell, IReadCell
    {
        public string ColumnName { get; set; }

        private string GetColumnName(MemberInfo member)
        {
            if (ColumnName.HasBlackSpace())
                return ColumnName;
            return member.Name;
        }

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
                if (headerData.HasLink)
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
            var linkDisplayText = $"@INDIRECT({displayTextAddress})";
            var linkFormula = $"HYPERLINK({linkedCellAddress},{linkDisplayText})";
            var fullFormula = optional ?
                $"=IFNA({linkFormula}, \"\")"
                :
                $"={linkFormula}";
            cellLink.SetFormulaA1(fullFormula);
            return cellLink;
        }

        public IXLCell WriteHeader<TResource>(IXLWorksheet worksheet, IHeaderData headerDataObj, Type[] otherTypes, TResource[] resources)
        {
            var headerData = (HeaderData)headerDataObj;
            var colIndex = headerData.ColumnsIndex;
            var member = headerData.MemberInfo;

            var value = GetColumnName(member);
            var cellValue = XLCellValue.FromObject(value);
            var cell = worksheet.Cell(1, colIndex);
            cell.SetValue(cellValue);

            if (headerData.HasLink)
            {
                var cellLink = worksheet.Cell(1, colIndex + 1);
                var cellLinkValue = XLCellValue.FromObject($"{headerData.externalSheetName} Link");
                cellLink.SetValue(cellLinkValue);
            }

            return cell;

            
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

        private interface IMatchThisHeaderData : IMatchHeaderData
        {
            bool IsMatch { get; }
        }

        private class NoMatchHeaderData : IMatchThisHeaderData
        {
            public bool IsMatch => false;
        }

        private class HeaderMatch : IMatchThisHeaderData
        {
            public bool IsMatch => true;

            public IXLCell[] relivantCells;

            public MemberInfo memberInfo;
        }

        public IMatchHeaderData MatchHeaders<TResource>(IXLRow headerRow, MemberInfo memberInfo)
        {
            var colName = GetColumnName(memberInfo);
            var relivantCells = headerRow.Cells()
                .Where(cell => cell.Value.IsText)
                .Where(cell =>
                    cell.Value.TryGetText(out var text)
                    && String.Equals(text, colName, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (relivantCells.None())
                return new NoMatchHeaderData();

            return new HeaderMatch
            {
                relivantCells = relivantCells,
                memberInfo = memberInfo,
            };
        }

        public TResource ReadCell<TResource>(TResource resource, IXLRow headerRow, IMatchHeaderData headerDataObj)
        {
            var headerData = (IMatchThisHeaderData)headerDataObj;
            if (!headerData.IsMatch)
                return resource;

            var headerMatchData = (HeaderMatch)headerData;
            var member = headerMatchData.memberInfo;

            var cell = headerRow.Cell(headerMatchData.relivantCells.First().Address.ColumnNumber);
            if (cell.Value.TryConvert(out Blank blankValue))
                return resource;


            var memberType = member.GetPropertyOrFieldType();
            return ProcessType(memberType);

            TResource ProcessType(Type typeToProcess)
            {
                return typeof(XLCellValue)
                    .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                    .Where(method => method.Name == nameof(XLCellValue.TryConvert))
                    .Where(
                        method =>
                        {
                            var paramsInQuestion = method.GetParameters();
                            if (paramsInQuestion.None())
                                return false;
                            var firstParam = paramsInQuestion.First();
                            var firstParamType = firstParam.ParameterType;
                            if (!firstParamType.IsByRef)
                                return false;
                            var paramTypeTocCheck = firstParamType.GetElementType();
                            var isMatch = paramTypeTocCheck == typeToProcess;
                            return isMatch;
                        })
                    .First(
                        (method, next) =>
                        {
                            if (!TryGetValue(out var valueToSet))
                                return resource;

                            var updatedResource = (TResource)member.SetPropertyOrFieldValue(resource, valueToSet);
                            return updatedResource;

                            bool TryGetValue(out object valueToSet)
                            {
                                var methodParams = method.GetParameters();
                                var parameters = methodParams.Length == 1 ?
                                    new object[] { null }
                                    :
                                    new object[] { null, System.Globalization.CultureInfo.CurrentCulture };

                                object result = method.Invoke(cell.Value, parameters);
                                var blResult = (bool)result;
                                valueToSet = parameters[0];
                                return blResult;
                            }
                        },
                        () =>
                        {
                            if (cell.Value.IsText)
                            {
                                if (cell.Value.GetText().TryParseRef(typeToProcess, out var parsedRefObj))
                                {
                                    var updatedResource = (TResource)member
                                        .SetPropertyOrFieldValue(resource, parsedRefObj);
                                    return updatedResource;
                                }
                            }
                            if (typeToProcess.IsAssignableFrom(typeof(string)))
                            {
                                var value = cell.GetText();
                                var updatedResource = (TResource)member
                                    .SetPropertyOrFieldValue(resource, value);
                                return updatedResource;
                            }
                            if (typeToProcess.IsAssignableFrom(typeof(int)))
                            {
                                if (cell.Value.IsNumber)
                                {
                                    var value = (int)cell.Value.GetNumber();
                                    var updatedResource = (TResource)member
                                        .SetPropertyOrFieldValue(resource, value);
                                    return updatedResource;
                                }
                                if (cell.Value.IsText)
                                {
                                    var valueStr = cell.Value.GetText();
                                    if (int.TryParse(valueStr, out int value))
                                    {
                                        var updatedResource = (TResource)member
                                            .SetPropertyOrFieldValue(resource, value);
                                        return updatedResource;
                                    }
                                }
                            }
                            if (typeToProcess.IsEnum)
                            {
                                var enumStr = cell.GetText();
                                if (Enum.TryParse(typeToProcess, enumStr, out var enumValue))
                                {
                                    var updatedResource = (TResource)member
                                        .SetPropertyOrFieldValue(resource, enumValue);
                                    return updatedResource;
                                }
                            }
                            return typeToProcess.IsNullable(
                                nullableBaseType =>
                                {
                                    var updatedResource = ProcessType(nullableBaseType);
                                    return updatedResource;
                                },
                                () =>
                                {
                                    return resource;
                                });
                        });
            }
        }
    }
}

