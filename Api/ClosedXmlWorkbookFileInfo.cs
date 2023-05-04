using System;
using System.Net.Mime;
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
    }
}

