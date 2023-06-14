using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper;
using EastFive.Linq;
using EastFive.Sheets.Api;

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
    }
}

