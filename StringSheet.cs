using System;
using System.Collections.Generic;
using System.Linq;

using EastFive;
using EastFive.Extensions;
using EastFive.Collections;
using EastFive.Linq;
using EastFive.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EastFive.Sheets
{
	public struct StringSheet
	{
        private string[][] data;

		public StringSheet(string[][] data)
		{
            this.data = data;
		}

        public static StringSheet CreateFromSheet(ISheet sheet, System.Text.Encoding[] encodingsToUse)
        {
            var data = sheet
                .ReadRows(encodingsToUse: encodingsToUse)
                .Where(row => row.Any(c => c.HasBlackSpace()))
                .Select(row => row.Select(c => c.Replace('\n', ' ').Trim()).ToArray())
                .ToArray();
            return new StringSheet(data);
        }

        public string[][] RawData => data;
        public string [] ColumnNames => data.First();

        public IEnumerable<(int row, int col, string value)> CellsAndIndexesWithHeaders => data
            .Select(
                (row, rowIndex) =>
                {
                    return row
                        .Select(
                            (cellValue, colIndex) =>
                            {
                                return (rowIndex, colIndex, cellValue);
                            })
                        .ToArray();
                })
            .SelectMany();

        public IEnumerable<string[]> RowsWithHeaders => data;
        public IEnumerable<string[]> RowsBody => data.Skip(1).ToArray();

        public int RowCount => data.Length;

        public StringSheet AppendColumnConstantValue(string header, string constantValue)
		{
            var newRows = data
                .Select(
                    (row, index) =>
                    {
                        var colValue = (index == 0) ? // If its the first row
                                header // Append header column name
                                :
                                constantValue; // Append constantValue

                        return row
                            .Append(colValue)
                            .ToArray();
                    })
                .ToArray();

            return new StringSheet(newRows);
        }

        public StringSheet AppendColumn(string header, Func<int, string[], string> populateRow)
        {
            var newRows = data
                .Select(
                    (row, index) =>
                    {
                        var colValue = (index == 0) ? // If its the first row
                                header // Append header column name
                                :
                                populateRow(index, row); // Append constantValue

                        return row
                            .Append(colValue)
                            .ToArray();
                    })
                .ToArray();

            return new StringSheet(newRows);

        }

        public StringSheet Truncate(int rowsToTake)
        {
            var newRows = data
                .Take(rowsToTake)
                .ToArray();

            return new StringSheet(newRows);
        }

        public StringSheet SanitizeCells(Func<string, string> callback)
        {
            var newRows = data
                .Select(
                    row =>
                    {
                        return row
                            .Select(
                                cell => callback(cell))
                            .ToArray();
                    })
                .ToArray();

            return new StringSheet(newRows);

        }

        public (string[][], StringSheet) StripFirstRows(int rowsToStrip)
        {
            var preamble = data.Take(rowsToStrip).ToArray();
            var newRows = data
                .Skip(rowsToStrip)
                .ToArray();

            return (preamble, new StringSheet(newRows));
        }

        public IEnumerable<(int row, int col, string cellValue, TItem mapping)> MapItemsToCells<TItem>(TItem[] matches,
            Func<TItem, (int row, int col, string value), bool> predicate)
        {
            return CellsAndIndexesWithHeaders
                .Select(
                    cellInfo =>
                    {
                        return matches
                            .Where(
                                match =>
                                {
                                    return predicate(match, cellInfo);
                                })
                            .First(
                                (match, next) =>
                                {
                                    return (true, cellInfo, match: match);
                                },
                                () =>
                                {
                                    return (false, cellInfo, match: default(TItem));
                                });
                    })
                .SelectWhere()
                .Select(tpl => (row: tpl.Item1.row, col: tpl.Item1.col, cellValue: tpl.Item1.value, mapping: tpl.Item2));
        }

        public IEnumerable<string> GetColumnValues(int index, string defaultValue)
        {
            return RowsBody
                .Select(
                    row =>
                    {
                        if(row.Length > index)
                            return row[index];
                        return defaultValue;
                    });
        }

        public IEnumerable<string> GetColumnValuesWithHeader(int index, string defaultValue)
        {
            return data
                .Select(
                    row =>
                    {
                        if (row.Length > index)
                            return row[index];
                        return defaultValue;
                    });
        }

        public IEnumerable<string> GetColumnValuesSkipShortRows(int index)
        {
            return RowsBody
                .Where(row => row.Length > index)
                .Select(row => row[index]);
        }

        public IEnumerable<string> UniqueValuesWhereHeader(Func<(string name, int index), bool> headerPredicate)
        {
            var mask = ColumnNames
                .Select(
                    (name, index) => (name, index))
                .Where(headerPredicate)
                .Select(tpl => tpl.index)
                .ToArray();
            return RowsBody
                .SelectMany(row => row.SelectIndexes(mask))
                .Distinct();
        }

        public bool TryGetHeaderIndex(string header, StringComparison comparison, out int index)
        {
            (var result, index) = ColumnNames
                .Select(
                    (header, index) => (header, index))
                .Where(tpl => String.Equals(tpl.header, header, comparison))
                .First(
                    (tpl, next) =>
                    {
                        return (true, tpl.index);
                    },
                    () =>
                    {
                        return (false, -1);
                    });
            return result;
        }

        public int GetRowLength(int index)
        {
            if (data.Length > index)
                return data[index].Length;
            return -1;
        }


    }
}

