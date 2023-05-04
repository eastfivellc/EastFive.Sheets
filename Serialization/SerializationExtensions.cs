using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
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
	}
}

