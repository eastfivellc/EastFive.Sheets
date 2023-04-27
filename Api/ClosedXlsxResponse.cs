using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;

using EastFive;
using EastFive.Api;
using EastFive.Extensions;
using EastFive.Sheets;
using EastFive.Serialization;
using EastFive.Collections;
using EastFive.Collections.Generic;
using EastFive.Linq;
using EastFive.Reflection;
using ClosedXML.Excel;

namespace EastFive.Api.Sheets
{
    [ClosedXlsxResponse]
    public delegate IHttpResponse ClosedXlsxResponse<TResource>(
            IDictionary<string, IEnumerable<TResource>> sheets,
            string filename = "");

    public class ClosedXlsxResponseAttribute : HttpGenericDelegateAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        [InstigateMethod]
        public IHttpResponse ContentResponse<TResource>(
            IDictionary<string, IEnumerable<TResource>> sheets,
            string filename = "")
        {
            var httpApiApp = this.httpApp as IApiApplication;
            var response = new ClosedXlsxResponse<TResource>(sheets,
                filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class ClosedXlsxResponse<T> : EastFive.Api.HttpResponse
        {
            private IDictionary<string, IEnumerable<T>> sheets;

            public ClosedXlsxResponse(
                    IDictionary<string, IEnumerable<T>> sheets,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(request, HttpStatusCode.OK)
            {
                // this.properties = properties;
                this.sheets = sheets;
                this.SetFileHeaders(
                    fileName.IsNullOrWhiteSpace() ? $"sheet.xlsx" : fileName,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                    true);
            }

            public override async Task WriteResponseAsync(Stream responseStream)
            {
                try
                {
                    var xlsData = ConvertToXlsxStreamAsync<T, byte[]>(sheets,
                        (data) => data);

                    await responseStream.WriteAsync(xlsData, 0, xlsData.Length,
                        this.Request.CancellationToken);
                }
                catch (Exception ex)
                {

                }
            }

            private static TResult ConvertToXlsxStreamAsync<TResource, TResult>(
                IDictionary<string, IEnumerable<T>> sheets,
                Func<byte[], TResult> callback)
            {
                using (var stream = new MemoryStream())
                {
                    var workbook = sheets
                        .Aggregate(
                            new XLWorkbook(),
                            (wb, sheet) =>
                            {
                                var ws = wb.Worksheets.Add(sheet.Key);
                                ws.Cell(1, 1).InsertData(sheet.Value);
                                return wb;
                            });
                    workbook.SaveAs(stream);

                    var buffer = stream.ToArray();
                    return callback(buffer);
                }
            }

        }
    }
}
