using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper;
using DocumentFormat.OpenXml.Spreadsheet;
using EastFive.Linq;
using EastFive.Reflection;
using EastFive.Serialization.Text;
using EastFive.Sheets.Api;
using Irony.Parsing;

namespace EastFive.Sheets
{
	public static class SerializationExtensions
	{
		public static IEnumerable<T> ReadSerializedValues<T>(this XLWorkbook workbook)
		{
			if (!typeof(T).TryGetAttributeInterface(out IReadSheet sheetReader))
				throw new Exception();

			var sheetName = sheetReader.GetSheetName(typeof(T));
			if (!workbook.TryGetWorksheet(sheetName, out IXLWorksheet sheet))
				return new T[] { };


            var rows = sheet
                .Rows();

            return sheetReader.ReadSheet<T>(sheet, rows);
		}

        public static async Task<Stream> WriteCSVAsync(this Stream streamToWriteTo, IEnumerable<string[]> csvData,
            bool leaveOpen = false)
        {
            using (TextWriter writer = new StreamWriter(streamToWriteTo, System.Text.Encoding.UTF8, leaveOpen: leaveOpen))
            {
                using (var csv = new CsvWriter(writer, System.Globalization.CultureInfo.InvariantCulture, leaveOpen: leaveOpen))
                {
                    foreach (var row in csvData)
                    {
                        foreach (var col in row)
                            csv.WriteConvertedField(col); // where values implements IEnumerable
                        await csv.NextRecordAsync();
                        await writer.FlushAsync();
                        await streamToWriteTo.FlushAsync();
                    }
                    return streamToWriteTo;
                }
            }
        }

        public static IEnumerable<TResource> ParseRows<TResource>(this IEnumerable<string[]> rows)
        {
            var parser = rows.GetEnumerator();

            if (!parser.MoveNext())
                yield break;

            var headers = parser.Current;
            var membersAndMappers = GetPropertyMappers<TResource>();

            while (parser.MoveNext())
            {
                TResource resource;
                try
                {
                    var fields = parser.Current;
                    var values = headers.CollateSimple(fields).ToArray();

                    resource = ParseResource(membersAndMappers, values);
                }
                catch (Exception ex)
                {
                    ex.GetType();
                    continue;
                }
                yield return resource;
            }

            TResource ParseResource(
                (MemberInfo, IMapTextProperty)[] membersAndMappers,
                (string key, string value)[] rowValues)
            {
                var resource = Activator.CreateInstance<TResource>();
                return membersAndMappers
                    .Aggregate(resource,
                        (resource, memberAndMapper) =>
                        {
                            var (member, mapper) = memberAndMapper;
                            return mapper.ParseRow(resource, member, rowValues);
                        });
            }
        }

        public static (MemberInfo, IMapTextProperty)[] GetPropertyMappers<TResource>()
        {
            return typeof(TResource)
                .GetPropertyOrFieldMembers()
                .Select(
                    member =>
                    {
                        var matchingAttrs = member
                            .GetAttributesInterface<IMapTextProperty>()
                            .ToArray();
                        return (member, matchingAttrs);
                    })
                .Where(tpl => tpl.matchingAttrs.Any())
                .Select(tpl => (tpl.member, attr: tpl.matchingAttrs.First()))
                .ToArray();
        }
    }
}

