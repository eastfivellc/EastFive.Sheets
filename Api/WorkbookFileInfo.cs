using System;
using System.Net.Mime;
using Microsoft.AspNetCore.Http;

namespace EastFive.Sheets.Api
{
	[WorkbookFileInfo]
	public class WorkbookFileInfo
	{
		public ContentDisposition contentDisposition;
		public ContentType contentType;
		public IHeaderDictionary headers;
		public long length;
		public IUnderstandSheets workbook;
    }
}

