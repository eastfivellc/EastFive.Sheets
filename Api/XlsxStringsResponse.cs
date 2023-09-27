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
using System.Reflection;

namespace EastFive.Api.Sheets
{
    [XlsxStringsResponse]
    public delegate IHttpResponse XlsxResponse(
            string[] headers,
            IEnumerable<string[]> resources,
            string filename = "",
            IDictionary<string, string> properties = default);

    public class XlsxStringsResponseAttribute : HttpDelegateAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        public override Task<IHttpResponse> InstigateInternal(IApplication httpApp, IHttpRequest request, ParameterInfo parameterInfo,
            Func<object, Task<IHttpResponse>> onSuccess)
        {
            XlsxResponse response = (headers, resources, filename, properties) =>
            {
                var httpApiApp = httpApp as IApiApplication;
                var response = new XlsxStringsResponse(headers, resources, properties,
                    filename,
                    httpApiApp, request);
                return UpdateResponse(parameterInfo, httpApp, request, response);
            };
            return onSuccess(response);
        }

        protected class XlsxStringsResponse : EastFive.Api.HttpResponse
        {
            private IDictionary<string, string> properties;
            private IEnumerable<string[]> resources;
            private string[] headers;

            public XlsxStringsResponse(string[] headers, 
                    IEnumerable<string[]> resources,
                    IDictionary<string, string> properties,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(request, HttpStatusCode.OK)
            {
                this.headers = headers;
                this.resources = resources;
                this.properties = properties;
                this.SetFileHeaders(
                    fileName.IsNullOrWhiteSpace() ? $"sheet.xlsx" : fileName,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                    true);
            }

            public override async Task WriteResponseAsync(Stream responseStream)
            {
                try
                {
                    var xlsData = ConvertToXlsxStreamAsync(headers, resources, properties,
                        (data) => data);

                    await responseStream.WriteAsync(xlsData, 0, xlsData.Length,
                        this.Request.CancellationToken);
                }
                catch (Exception ex)
                {
                    var msgBytes = ex.Message.GetBytes(Encoding.UTF8);
                    await responseStream.WriteAsync(msgBytes, 0, msgBytes.Length,
                        this.Request.CancellationToken);
                }
            }

            private static TResult ConvertToXlsxStreamAsync<TResult>(
                string [] headers, IEnumerable<string[]> resources, IDictionary<string, string> properties,
                Func<byte[], TResult> callback)
            {
                using (var stream = new MemoryStream())
                {
                    OpenXmlWorkbook.Create(stream,
                        (workbook) =>
                        {
                            #region Custom properties

                            workbook.WriteCustomProperties(
                                    (writeProp) =>
                                    {
                                        properties
                                            .NullToEmpty()
                                            .Select(
                                                prop =>
                                                {
                                                    writeProp(prop.Key, prop.Value);
                                                    return true;
                                                })
                                            .ToArray();
                                        return true;
                                    });

                            #endregion

                            bool didWrite = workbook.WriteSheetByRow(
                                (writeRow) =>
                                {
                                    #region Header 

                                    writeRow(headers);

                                    #endregion

                                    #region Body

                                    bool[] writeSuccesses = resources
                                        .Select(
                                            (rowValues, index) =>
                                            {
                                                writeRow(rowValues);
                                                return true;
                                            })
                                        .ToArray();

                                    #endregion

                                    return true;
                                });

                            return true;
                        });

                    var buffer = stream.ToArray();
                    return callback(buffer);
                }
            }

        }
    }
}
