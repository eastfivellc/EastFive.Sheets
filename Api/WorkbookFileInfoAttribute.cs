﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using Microsoft.AspNetCore.Http;

using EastFive.Api;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EastFive.Sheets.Api
{
    public class WorkbookFileInfoAttribute : Attribute,
        IProvideBlobValue
    {
        public TResult ProvideValue<TResult>(MultipartContentTokenParser valueToBind,
            Func<object, TResult> onBound,
            Func<string, TResult> onFailure)
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

        public TResult ProvideValue<TResult>(IFormFile valueToBind,
            Func<object, TResult> onBound,
            Func<string, TResult> onFailure)
        {
            var raw = valueToBind.OpenReadStream().ToBytes();

            return GetWorkbook(
                sheetUnderstander =>
                {

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
                },
                onFailure: onFailure);

            TResult GetWorkbook(
                Func< IUnderstandSheets, TResult> onBound,
                Func<string, TResult> onFailure)
            {
                if (valueToBind.ContentType.Contains("csv", StringComparison.OrdinalIgnoreCase))
                {
                    var stream = new MemoryStream(raw);
                    var workbook = new CsvWorkbook(stream);
                    var sheetUnderstander = (IUnderstandSheets)workbook;
                    return onBound(sheetUnderstander);
                }

                if (IsXlsx())
                {
                    var stream = new MemoryStream(raw);
                    var workbook = new OpenXmlWorkbook(stream);
                    var sheetUnderstander = (IUnderstandSheets)workbook;
                    return onBound(sheetUnderstander);
                }

                return onFailure($"Could not process file of type:`{valueToBind.ContentType}`");

                bool IsXlsx()
                {
                    var contentType = valueToBind.ContentType;

                    if (contentType.Contains("openxmlformats", StringComparison.OrdinalIgnoreCase))
                        if (contentType.Contains("sheet", StringComparison.OrdinalIgnoreCase))
                            return true;

                    if (contentType.Contains("xlsx", StringComparison.OrdinalIgnoreCase))
                        return true;

                    if (contentType.Contains("officedocument", StringComparison.OrdinalIgnoreCase))
                        if (contentType.Contains("sheet", StringComparison.OrdinalIgnoreCase))
                            return true;


                    return false;
                }
            }
        }
    }
}
