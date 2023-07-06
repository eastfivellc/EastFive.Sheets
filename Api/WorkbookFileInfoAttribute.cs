using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using EastFive.Api;
using EllipticCurve.Utils;
using Microsoft.AspNetCore.Http;

namespace EastFive.Sheets.Api
{
    public class WorkbookFileInfoAttribute : Attribute,
        IProvideBlobValue
    {
        public TResult ProvideValue<TResult>(MultipartContentTokenParser valueToBind, Func<object, TResult> onBound, Func<string, TResult> onFailure)
        {
            var raw = valueToBind.ReadStream().ToBytes();
            var stream = new MemoryStream(raw);
            var workbook = new OpenXmlWorkbook(stream);
            var sheetUnderstander = (IUnderstandSheets)workbook;
            var info = new WorkbookFileInfo
            {
                workbook = sheetUnderstander,
                raw = raw,
            };
            return onBound(info);
        }

        public TResult ProvideValue<TResult>(IFormFile valueToBind, Func<object, TResult> onBound, Func<string, TResult> onFailure)
        {
            var raw = valueToBind.OpenReadStream().ToBytes();
            var stream = new MemoryStream(raw);
            var workbook = new OpenXmlWorkbook(stream);
            var sheetUnderstander = (IUnderstandSheets)workbook;
            var info = new WorkbookFileInfo
            {
                contentDisposition = new ContentDisposition(valueToBind.ContentDisposition),
                contentType = new ContentType(valueToBind.ContentType),
                headers = valueToBind.Headers,
                length = valueToBind.Length,
                workbook = sheetUnderstander,
                raw = raw,
            };
            return onBound(info);
        }
    }
}
