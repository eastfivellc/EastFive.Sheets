using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using EastFive.Api;
using Microsoft.AspNetCore.Http;

namespace EastFive.Sheets.Api
{
    public class ClosedXmlWorkbookFileInfoAttribute : Attribute,
        IProvideBlobValue
    {
        public TResult ProvideValue<TResult>(MultipartContentTokenParser valueToBind, Func<object, TResult> onBound, Func<string, TResult> onFailure)
        {
            var stream = valueToBind.ReadStream();
            var workbook = new XLWorkbook(stream);
            var info = new ClosedXmlWorkbookFileInfo
            {
                workbook = workbook,
            };
            return onBound(info);
        }

        public TResult ProvideValue<TResult>(IFormFile valueToBind, Func<object, TResult> onBound, Func<string, TResult> onFailure)
        {
            var stream = valueToBind.OpenReadStream().ToCachedStream();
            var workbook = new XLWorkbook(stream);
            var info = new ClosedXmlWorkbookFileInfo
            {
                contentDisposition = new ContentDisposition(valueToBind.ContentDisposition),
                contentType = new ContentType(valueToBind.ContentType),
                headers = valueToBind.Headers,
                length = valueToBind.Length,
                workbook = workbook,
            };
            return onBound(info);
        }
    }
}
