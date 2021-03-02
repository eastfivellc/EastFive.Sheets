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

namespace EastFive.Api.Sheets
{
    [XlsxResponse]
    public delegate IHttpResponse XlsxResponse<TResource>(
            IDictionary<string, string> properties,
            IEnumerable<TResource> resources,
            string filename = "");

    public class XlsxResponseAttribute : HttpGenericDelegateAttribute
    {
        public override HttpStatusCode StatusCode => HttpStatusCode.OK;

        public override string Example => "<xml></xml>";

        [InstigateMethod]
        public IHttpResponse ContentResponse<TResource>(IDictionary<string, string> properties,
            IEnumerable<TResource> resources,
            string filename = "")
        {
            var httpApiApp = this.httpApp as IApiApplication;
            var response = new XlsxResponse<TResource>(properties, resources,
                filename,
                httpApiApp, this.request);
            return UpdateResponse(parameterInfo, httpApp, request, response);
        }

        protected class XlsxResponse<T> : EastFive.Api.HttpResponse
        {
            private IDictionary<string, string> properties;
            private IEnumerable<T> resources;

            public XlsxResponse(IDictionary<string, string> properties,
                    IEnumerable<T> resources,
                    string fileName,
                    IApiApplication httpApiApp,
                    IHttpRequest request)
                : base(request, HttpStatusCode.OK)
            {
                this.properties = properties;
                this.resources = resources;
                this.SetFileHeaders(
                    fileName.IsNullOrWhiteSpace() ? $"sheet.xlsx" : fileName,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                    true);
            }

            public override async Task WriteResponseAsync(Stream responseStream)
            {
                try
                {
                    var xlsData = ConvertToXlsxStreamAsync(properties, resources,
                        (data) => data);

                    await responseStream.WriteAsync(xlsData, 0, xlsData.Length,
                        this.Request.CancellationToken);
                }
                catch (Exception ex)
                {

                }
            }

            private static TResult ConvertToXlsxStreamAsync<TResource, TResult>(
                IDictionary<string, string> properties, IEnumerable<TResource> resources,
                Func<byte[], TResult> callback)
            {
                var guidReferences = resources
                    .Select(
                        (obj, index) =>
                        {
                            if (typeof(IReferenceable).IsInstanceOfType(obj))
                            {
                                var resourceId = (obj as IReferenceable).id;
                                return resourceId.PairWithValue($"A{index}");
                            }
                            return default(KeyValuePair<Guid, string>?);
                        })
                    .SelectWhereHasValue()
                    .ToDictionary();

                using (var stream = new MemoryStream())
                {
                    OpenXmlWorkbook.Create(stream,
                        (workbook) =>
                        {
                            #region Custom properties

                            workbook.WriteCustomProperties(
                                    (writeProp) =>
                                    {
                                        properties.Select(
                                            prop =>
                                            {
                                                writeProp(prop.Key, prop.Value);
                                                return true;
                                            }).ToArray();
                                        return true;
                                    });

                            #endregion

                            workbook.WriteSheetByRow(
                                    (writeRow) =>
                                    {
                                        var propertyOrder = typeof(TResource).GetPropertyOrFieldMembers();
                                        if (!propertyOrder.Any() && resources.Any())
                                        {
                                            propertyOrder = resources.First().GetType().GetProperties();
                                        }

                                    #region Header 

                                    var headers = propertyOrder
                                            .Select(
                                                propInfo => propInfo.GetCustomAttribute<JsonPropertyAttribute, string>(
                                                    attr => attr.PropertyName,
                                                    () => propInfo.Name))
                                            .ToArray();
                                        writeRow(headers);

                                    #endregion

                                    #region Body

                                    var rows = resources.Select(
                                            (result, index) =>
                                            {
                                                var values = propertyOrder
                                                    .Select(
                                                        property => property
                                                            .GetPropertyOrFieldValue(result)
                                                            .CastToXlsSerialization(property, guidReferences))
                                                    .ToArray();
                                                writeRow(values);
                                                return true;
                                            }).ToArray();

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
