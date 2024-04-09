using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Text;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;

namespace EastFive.Sheets.Api
{
	[ClosedXmlWorkbookFileInfo]
	public class ClosedXmlWorkbookFileInfo
    {
		public ContentDisposition contentDisposition;
		public ContentType contentType;
		public IHeaderDictionary headers;
		public long length;
		public XLWorkbook workbook;
		public IUnderstandSheets sheets
		{
			get
			{
                return new ClosedXmlWorkbook(this.workbook);
			}
		}

        private class ClosedXmlWorkbook : IUnderstandSheets
        {
            private XLWorkbook workbook;

            public ClosedXmlWorkbook(XLWorkbook workbook)
            {
                this.workbook = workbook;
            }

            public void Dispose()
            {
                this.workbook.Dispose();
            }

            public TResult ReadCustomValues<TResult>(Func<KeyValuePair<string, string>[], TResult> onResults)
            {
                throw new NotImplementedException();
            }

            public IEnumerable<ISheet> ReadSheets()
            {
                return workbook.Worksheets
                    .Select(
                        ws => (ISheet)new ClosedXmlWorkbookSheet(ws))
                    .ToArray();
            }

            public TResult WriteCustomProperties<TResult>(Func<Action<string, string>, TResult> callback)
            {
                throw new NotImplementedException();
            }

            public TResult WriteSheetByRow<TResult>(Func<Action<object[]>, TResult> p, string sheetName = "Sheet1")
            {
                throw new NotImplementedException();
            }
        }

        private class ClosedXmlWorkbookSheet : ISheet
        {
            private IXLWorksheet ws;

            public ClosedXmlWorkbookSheet(IXLWorksheet ws)
            {
                this.ws = ws;
            }

            public string Name => ws.Name;

            public IEnumerable<string[]> ReadRows(Func<Type, object,
                Func<string>, string> discardSerializer = default,
                bool discardAutoDecodeEncoding = default,
                Encoding[] discardEncodingsToUse = default)
            {
                return ws.Rows()
                    .Select(
                        row =>
                        {
                            return row.Cells()
                                .Select(
                                    cell =>
                                    {
                                        return cell.GetText();
                                    })
                                .ToArray();
                        })
                    .ToArray();
            }

            public void WriteRows(string fileName, object[] rows)
            {
                throw new NotImplementedException();
            }
        }
    }

}

