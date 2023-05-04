using System;
using EastFive.Api;
using EastFive.Api.Sheets;
using EastFive.Reflection;

using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using ClosedXML.Excel;
using EastFive.Extensions;
using System.Linq;

namespace EastFive.Sheets.Api
{
    [WorkbookResponse]
    public delegate IHttpResponse WorkbookResponse<TResource1, TResource2, TResource3, TResource4>(
            TResource1[] resource1s, TResource2[] resource2s, TResource3[] resource3s, TResource4[] resource4s,
            string filename = "");

    public class WorkbookResponseAttribute : HttpGenericDelegateAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        [InstigateMethod]
        public IHttpResponse ContentResponse<TResource1, TResource2, TResource3, TResource4>(
            TResource1[] resource1s, TResource2[] resource2s, TResource3[] resource3s, TResource4[] resource4s,
            string filename = "")
        {
            var httpApiApp = this.httpApp as IApiApplication;
            var response = new WorkbookResponse<TResource1, TResource2, TResource3, TResource4>(
                resource1s:resource1s, resource2s:resource2s, resource3s:resource3s, resource4s:resource4s,
                filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class WorkbookResponse<TResource1, TResource2, TResource3, TResource4> : EastFive.Api.HttpResponse
        {
            TResource1[] resource1s; TResource2[] resource2s;
            TResource3[] resource3s; TResource4[] resource4s;

            public WorkbookResponse(
                    TResource1[] resource1s, TResource2[] resource2s,
                    TResource3[] resource3s, TResource4[] resource4s,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(request, HttpStatusCode.OK)
            {
                this.resource1s = resource1s;
                this.resource2s = resource2s;
                this.resource3s = resource3s;
                this.resource4s = resource4s;

                this.SetFileHeaders(
                    fileName.IsNullOrWhiteSpace() ? $"sheet.xlsx" : fileName,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                    true);
            }

            public override async Task WriteResponseAsync(Stream responseStream)
            {
                try
                {
                    var xlsData = ConvertToXlsxStreamAsync(
                        resource1s: resource1s, resource2s: resource2s,
                        resource3s: resource3s, resource4s: resource4s);

                    await responseStream.WriteAsync(xlsData, 0, xlsData.Length,
                        this.Request.CancellationToken);
                }
                catch (Exception ex)
                {

                }
            }

            private byte[] ConvertToXlsxStreamAsync(
                    TResource1[] resource1s, TResource2[] resource2s,
                    TResource3[] resource3s, TResource4[] resource4s)
            {
                using (var stream = new MemoryStream())
                {
                    var wb = new XLWorkbook();
                    WriteSheet<TResource1>(resource1s, new Type[] { typeof(TResource2), typeof(TResource3), typeof(TResource4) });
                    WriteSheet<TResource2>(resource2s, new Type[] { typeof(TResource1), typeof(TResource3), typeof(TResource4) });
                    WriteSheet<TResource3>(resource3s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource4) });
                    WriteSheet<TResource4>(resource4s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource3) });
                    wb.SaveAs(stream);

                    var buffer = stream.ToArray();
                    return buffer;

                    IWriteSheet GetSheetWriter(Type type)
                    {
                        if (!type.TryGetAttributeInterface(out IWriteSheet sheetWriter))
                            throw new Exception($"Cannot return objects of type {type.FullName} from {this.GetType().FullName}"
                                + $" without attribute implementing {nameof(IWriteSheet)}.");
                        return sheetWriter;
                    }

                    void WriteSheet<TResource>(TResource[] resources, Type[] otherTypes)
                    {
                        if (resources.IsDefaultOrNull())
                            return;

                        var sheetWriter = GetSheetWriter(typeof(TResource));
                        var sheetName = sheetWriter.GetSheetName(typeof(TResource));
                        var ws = wb.Worksheets.Add(sheetName);
                        var headerData = sheetWriter.GetHeaderData(typeof(TResource), otherTypes);
                        sheetWriter.WriteSheet(ws, resources, otherTypes, headerData);

                    }
                }
            }

        }
    }
}

