using System;
using EastFive.Api;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using ClosedXML.Excel;
using EastFive.Extensions;
using EastFive.Serialization;

namespace EastFive.Sheets.Api
{
    [WorkbookResponse1]
    public delegate IHttpResponse WorkbookResponse<TResource1>(
            TResource1[] resource1s, string filename = "");

    [WorkbookResponse2]
    public delegate IHttpResponse WorkbookResponse<TResource1, TResource2>(
            TResource1[] resource1s, TResource2[] resource2s, string filename = "");

    [WorkbookResponse4]
    public delegate IHttpResponse WorkbookResponse<TResource1, TResource2, TResource3, TResource4>(
            TResource1[] resource1s, TResource2[] resource2s, TResource3[] resource3s, TResource4[] resource4s,
            string filename = "");

    [WorkbookResponse5]
    public delegate IHttpResponse WorkbookResponse<TResource1, TResource2, TResource3, TResource4, TResource5>(
            TResource1[] resource1s, TResource2[] resource2s, TResource3[] resource3s, TResource4[] resource4s, TResource5[] resource5s,
            string filename = "");

    public abstract class WorkbookResponseAttribute : HttpGenericDelegateAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        protected abstract class WorkbookResponse : EastFive.Api.HttpResponse
        {
            public WorkbookResponse(
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(request, HttpStatusCode.OK)
            {
                this.SetFileHeaders(
                    fileName.IsNullOrWhiteSpace() ? $"sheet.xlsx" : fileName,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                    true);
            }

            public override async Task WriteResponseAsync(Stream responseStream)
            {
                try
                {
                    var xlsData = ConvertToXlsxStreamAsync();

                    await responseStream.WriteAsync(xlsData, 0, xlsData.Length,
                        this.Request.CancellationToken);
                }
                catch (Exception ex)
                {
                    var msgBytes = ex.Message.GetBytes(System.Text.Encoding.UTF8);
                    await responseStream.WriteAsync(msgBytes, 0, msgBytes.Length,
                        this.Request.CancellationToken);
                }
            }

            private byte[] ConvertToXlsxStreamAsync()
            {
                using (var stream = new MemoryStream())
                {
                    var wb = new XLWorkbook();
                    WriteSheets(wb);
                    wb.SaveAs(stream);

                    var buffer = stream.ToArray();
                    return buffer;
                }
            }

            protected void WriteSheet<TResource>(XLWorkbook wb, TResource[] resources, Type[] otherTypes)
            {
                if (resources.IsDefaultOrNull())
                    return;

                var sheetWriter = GetSheetWriter(typeof(TResource));
                var sheetName = sheetWriter.GetSheetName(typeof(TResource));
                var ws = wb.Worksheets.Add(sheetName);
                var headerData = sheetWriter.GetHeaderData(typeof(TResource), otherTypes);
                sheetWriter.WriteSheet(ws, resources, otherTypes, headerData);

            }

            IWriteSheet GetSheetWriter(Type type)
            {
                if (!type.TryGetAttributeInterface(out IWriteSheet sheetWriter))
                    throw new Exception($"Cannot return objects of type {type.FullName} from {this.GetType().FullName}"
                        + $" without attribute implementing {nameof(IWriteSheet)}.");
                return sheetWriter;
            }

            protected abstract void WriteSheets(XLWorkbook wb);
        }
    }

    public class WorkbookResponse1Attribute : WorkbookResponseAttribute
    {
        [InstigateMethod]
        public IHttpResponse ContentResponse<TResource1>(TResource1[] resource1s, string filename = "")
        {
            var httpApiApp = this.httpApp as IApiApplication;
            var response = new WorkbookResponse<TResource1>(
                resource1s: resource1s, filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class WorkbookResponse<TResource1> : WorkbookResponse
        {
            TResource1[] resource1s;

            public WorkbookResponse(
                    TResource1[] resource1s,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(fileName, httpApiApp, request)
            {
                this.resource1s = resource1s;
            }

            protected override void WriteSheets(XLWorkbook wb)
            {
                WriteSheet<TResource1>(wb, resource1s, new Type[] { });
            }
        }
    }

    public class WorkbookResponse2Attribute : WorkbookResponseAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        [InstigateMethod]
        public IHttpResponse ContentResponse<TResource1, TResource2>(
            TResource1[] resource1s, TResource2[] resource2s,
            string filename = "")
        {
            var httpApiApp = this.httpApp as IApiApplication;
            var response = new WorkbookResponse<TResource1, TResource2>(
                resource1s:resource1s, resource2s:resource2s,
                filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class WorkbookResponse<TResource1, TResource2> : WorkbookResponse
        {
            TResource1[] resource1s; TResource2[] resource2s;

            public WorkbookResponse(
                    TResource1[] resource1s, TResource2[] resource2s,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(fileName, httpApiApp, request)
            {
                this.resource1s = resource1s;
                this.resource2s = resource2s;
            }

            protected override void WriteSheets(XLWorkbook wb)
            {
                WriteSheet<TResource1>(wb, resource1s, new Type[] { typeof(TResource2)});
                WriteSheet<TResource2>(wb, resource2s, new Type[] { typeof(TResource1)});
            }
        }
    }

    public class WorkbookResponse4Attribute : WorkbookResponseAttribute
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
                resource1s: resource1s, resource2s: resource2s, resource3s: resource3s, resource4s: resource4s,
                filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class WorkbookResponse<TResource1, TResource2, TResource3, TResource4> : WorkbookResponse
        {
            TResource1[] resource1s; TResource2[] resource2s;
            TResource3[] resource3s; TResource4[] resource4s;

            public WorkbookResponse(
                    TResource1[] resource1s, TResource2[] resource2s,
                    TResource3[] resource3s, TResource4[] resource4s,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(fileName, httpApiApp, request)
            {
                this.resource1s = resource1s;
                this.resource2s = resource2s;
                this.resource3s = resource3s;
                this.resource4s = resource4s;
            }

            protected override void WriteSheets(XLWorkbook wb)
            {
                WriteSheet<TResource1>(wb, resource1s, new Type[] { typeof(TResource2), typeof(TResource3), typeof(TResource4) });
                WriteSheet<TResource2>(wb, resource2s, new Type[] { typeof(TResource1), typeof(TResource3), typeof(TResource4) });
                WriteSheet<TResource3>(wb, resource3s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource4) });
                WriteSheet<TResource4>(wb, resource4s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource3) });
            }
        }
    }

    public class WorkbookResponse5Attribute : WorkbookResponseAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        [InstigateMethod]
        public IHttpResponse ContentResponse<TResource1, TResource2, TResource3, TResource4, TResource5>(
            TResource1[] resource1s, TResource2[] resource2s, TResource3[] resource3s, TResource4[] resource4s, TResource5[] resource5s,
            string filename = "")
        {
            var httpApiApp = this.httpApp as IApiApplication;
            var response = new WorkbookResponse<TResource1, TResource2, TResource3, TResource4, TResource5>(
                resource1s: resource1s, resource2s: resource2s, resource3s: resource3s, resource4s: resource4s, resource5s: resource5s,
                filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class WorkbookResponse<TResource1, TResource2, TResource3, TResource4, TResource5> : WorkbookResponse
        {
            TResource1[] resource1s; TResource2[] resource2s;
            TResource3[] resource3s; TResource4[] resource4s;
            TResource5[] resource5s;

            public WorkbookResponse(
                    TResource1[] resource1s, TResource2[] resource2s,
                    TResource3[] resource3s, TResource4[] resource4s,
                    TResource5[] resource5s,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(fileName, httpApiApp, request)
            {
                this.resource1s = resource1s;
                this.resource2s = resource2s;
                this.resource3s = resource3s;
                this.resource4s = resource4s;
                this.resource5s = resource5s;
            }

            protected override void WriteSheets(XLWorkbook wb)
            {
                WriteSheet<TResource1>(wb, resource1s, new Type[] { typeof(TResource2), typeof(TResource3), typeof(TResource4), typeof(TResource5) });
                WriteSheet<TResource2>(wb, resource2s, new Type[] { typeof(TResource1), typeof(TResource3), typeof(TResource4), typeof(TResource5) });
                WriteSheet<TResource3>(wb, resource3s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource4), typeof(TResource5) });
                WriteSheet<TResource4>(wb, resource4s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource3), typeof(TResource5) });
                WriteSheet<TResource5>(wb, resource5s, new Type[] { typeof(TResource1), typeof(TResource2), typeof(TResource3), typeof(TResource4) });
            }
        }
    }
}

