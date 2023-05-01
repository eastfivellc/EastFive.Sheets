using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using EastFive.Api;
using Microsoft.AspNetCore.Http;

namespace EastFive.Sheets
{
    public class WorkbookApiAttribute : Attribute,
        IProvideBlobValue
    {
        public TResult ProvideValue<TResult>(MultipartContentTokenParser valueToBind, Func<object, TResult> onBound, Func<string, TResult> onFailure)
        {
            var stream = valueToBind.ReadStream();
            var workbook = new OpenXmlWorkbook(stream);
            var sheetUnderstander = (IUnderstandSheets)workbook;
            return onBound(sheetUnderstander);
        }

        public TResult ProvideValue<TResult>(IFormFile valueToBind,
            Func<object, TResult> onBound,
            Func<string, TResult> onFailure)
        {
            var stream = valueToBind.OpenReadStream();
            var workbook = new OpenXmlWorkbook(stream);
            var sheetUnderstander = (IUnderstandSheets)workbook;
            return onBound(sheetUnderstander);
        }
    }
}
