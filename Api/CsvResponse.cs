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
using CsvHelper;
using System.Reflection;

namespace EastFive.Api.Sheets
{
    [CsvResponse]
    public delegate IHttpResponse CsvResponse(
            IEnumerable<string[]> values,
            string filename = "");

    public class CsvResponseAttribute : HttpDelegateAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        public override Task<IHttpResponse> InstigateInternal(IApplication httpApp, IHttpRequest request, ParameterInfo parameterInfo,
            Func<object, Task<IHttpResponse>> onSuccess)
        {
            CsvResponse response = (values, filename) =>
            {
                var httpApiApp = httpApp as IApiApplication;
                var response = new HttpCsvResponse(values,
                    filename,
                    httpApiApp, request);
                return UpdateResponse(parameterInfo, httpApp, request, response);
            };
            return onSuccess(response);
        }

        protected class HttpCsvResponse : EastFive.Api.HttpResponse
        {
            private IEnumerable<string[]> values;

            public HttpCsvResponse(IEnumerable<string[]> values,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(request, HttpStatusCode.OK)
            {
                this.values = values;
                this.SetFileHeaders(
                    fileName.IsNullOrWhiteSpace() ? $"sheet.csv" : fileName,
                    "text/csv",
                    true);
            }

            public override async Task WriteResponseAsync(Stream responseStream)
            {
                try
                {
                    await responseStream.WriteCSVAsync(values);
                }
                catch (Exception ex)
                {
                    var msgBytes = ex.Message.GetBytes(Encoding.UTF8);
                    await responseStream.WriteAsync(msgBytes, 0, msgBytes.Length,
                        this.Request.CancellationToken);
                }
            }

        }
    }
}
