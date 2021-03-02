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
using System.Net.Http;
using EastFive.Linq.Async;

namespace EastFive.Api.Sheets
{
    public static class XlsxResponseExtensions
    {

        #region Xlsx

        //public static IHttpResponse CreateXlsxResponse(this IHttpRequest request, Stream xlsxData, string filename = "")
        //{
        //    return CreateXlsxResponse(request, xlsxData)
        //    var response = new StringHttpResponse(request, HttpStatusCode.OK,
        //        String.IsNullOrWhiteSpace(filename) ? $"sheet.xlsx" : filename,
        //        "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
        //        false,
        //        xlsxContent,
        //        default);
        //    return response;
        //}

        public static IHttpResponse CreateXlsxResponse(this IHttpRequest request,
            byte[] xlsxData, string filename = "")
        {
            var response = new BytesHttpResponse(request, HttpStatusCode.OK,
                filename.IsNullOrWhiteSpace() ? $"sheet.xlsx" : filename,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                false,
                xlsxData);
            return response;
        }

        public static IHttpResponse CreateXlsxResponse<TResource>(this IHttpRequest request,
            IDictionary<string, string> properties, IEnumerable<TResource> resources,
            string filename = "")
        {
            try
            {
                var responseStream = ConvertToXlsxStreamAsync(properties, resources,
                    (stream) => stream);
                var response = request.CreateXlsxResponse(responseStream, filename);
                return response;
            }
            catch (Exception ex)
            {
                return request
                    .CreateResponse(HttpStatusCode.Conflict, ex.StackTrace)
                    //.AddReason(ex.Message);
                    ;
            }
        }

        public static IHttpResponse CreateMultisheetXlsxResponse<TResource>(this IHttpRequest request,
            IDictionary<string, string> properties, IEnumerable<TResource> resources,
            string filename = "")
            where TResource : IReferenceable
        {
            try
            {
                var responseStream = ConvertToMultisheetXlsxStreamAsync(properties, resources,
                    (stream) => stream);
                var response = request.CreateXlsxResponse(responseStream, filename);
                return response;
            }
            catch (Exception ex)
            {
                return request.CreateResponse(HttpStatusCode.Conflict, ex.StackTrace);
                // .AddReason(ex.Message);
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
                                var propertyOrder = typeof(TResource).GetProperties();
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
                                                property => property.GetValue(result).CastToXlsSerialization(property, guidReferences))
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

        private static TResult ConvertToMultisheetXlsxStreamAsync<TResource, TResult>(
            IDictionary<string, string> properties, IEnumerable<TResource> resources,
            Func<byte[], TResult> callback)
            where TResource : IReferenceable
        {
            var resourceGroups = resources
                .GroupBy(
                    (res) =>
                    {
                        var resourceId = res.id;
                        return typeof(TResource).Name;
                    });

            var guidReferences = resourceGroups
                .SelectMany(
                    grp =>
                        grp
                            .Select(
                                (res, index) =>
                                {
                                    var resourceId = res.id;
                                    return resourceId.PairWithValue($"{grp.Key}!A{index + 2}");
                                }))
                .ToDictionary();

            using (var stream = new MemoryStream())
            {
                bool wroteRows = OpenXmlWorkbook.Create(stream,
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

                        foreach (var resourceGrp in resourceGroups)
                        {
                            bool wroteRow = workbook.WriteSheetByRow(
                                (writeRow) =>
                                {
                                    var resourcesForSheet = resourceGrp.ToArray();
                                    if (!resourcesForSheet.Any())
                                        return false;
                                    var resource = resourcesForSheet.First();
                                    if (resource.IsDefault())
                                        return false;
                                    var propertyOrder = resource.GetType().GetProperties().Reverse().ToArray();

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

                                    var rows = resourcesForSheet.Select(
                                        (result, index) =>
                                        {
                                            var values = propertyOrder
                                            .Select(
                                                property => property.GetValue(result).CastToXlsSerialization(property, guidReferences))
                                            .ToArray();
                                            writeRow(values);
                                            return true;
                                        }).ToArray();

                                    #endregion

                                    return true;
                                },
                                resourceGrp.Key);
                        }

                        return true;
                    });

                stream.Flush();
                var buffer = stream.ToArray();
                return callback(buffer);
            }
        }

        public static object CastToXlsSerialization(this object obj,
            MemberInfo property, IDictionary<Guid, string> lookups)
        {
            if (obj is IReferenceable)
            {
                var webId = obj as IReferenceable;
                var objId = webId.id;

                var resourceDisplayValue = obj.GetType().Name;
                if (property.Name == "Id" || !lookups.ContainsKey(objId)) // TODO: Use custom property attributes
                    return resourceDisplayValue;

                return new OpenXmlWorkbook.CellReference
                {
                    value = resourceDisplayValue,
                    formula = lookups[objId],
                };
            }
            return obj;
        }

        private static TResult TryCastFromXlsSerialization<TResult>(PropertyInfo property, string valueString,
            Func<object, TResult> onParsed,
            Func<TResult> onNotParsed)
        {
            if (property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
                return TryCastFromXlsSerialization(propertyType, valueString,
                    (value) =>
                    {
                        dynamic objValue = System.Activator.CreateInstance(property.PropertyType);
                        objValue = value;
                        return onParsed((object)objValue);
                    },
                    onNotParsed);
            }

            if (property.PropertyType.IsEnum)
            {
                var enumUnderlyingType = Enum.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
                if (int.TryParse(valueString, out int enumValueInt))
                {
                    object underlyingValueInt = System.Convert.ChangeType(enumValueInt, enumUnderlyingType);
                    return onParsed(underlyingValueInt);
                }
                var values = Enum.GetValues(property.PropertyType);
                var names = Enum.GetNames(property.PropertyType);
                if (!names.Contains(valueString))
                    return onNotParsed();

                return names.IndexOf(valueString, (s1, s2) => String.Compare(s1, s2, true) == 0,
                    (valueIndex) =>
                    {
                        var enumValue = values.GetValue(valueIndex);
                        object underlyingValue = System.Convert.ChangeType(enumValue, enumUnderlyingType);
                        return onParsed(underlyingValue);
                    },
                    onNotParsed);
            }

            return TryCastFromXlsSerialization(property.PropertyType, valueString,
                    onParsed, onNotParsed);
        }

        private static TResult TryCastFromXlsSerialization<TResult>(Type type, string valueString,
            Func<object, TResult> onParsed,
            Func<TResult> onNotParsed)
        {
            //if (type.GUID == typeof(WebId).GUID)
            //    if (Uri.TryCreate(valueString, UriKind.Absolute, out Uri urn))
            //        return urn.TryParseWebUrn(
            //            (nid, ns, uuid) => onParsed(new WebId
            //            {
            //                UUID = uuid,
            //            }),
            //            (why) => onNotParsed());
            //    else
            //        return onNotParsed();

            if (type.GUID == typeof(string).GUID)
                return onParsed(valueString);

            return type.TryParse<bool, TResult>(valueString, bool.TryParse, onParsed,
                () => type.TryParse<Guid, TResult>(valueString, Guid.TryParse, onParsed,
                () => type.TryParse<Int32, TResult>(valueString, Int32.TryParse, onParsed,
                () => type.TryParse<decimal, TResult>(valueString, decimal.TryParse, onParsed,
                () => type.TryParse<double, TResult>(valueString, double.TryParse, onParsed,
                () => type.TryParse<float, TResult>(valueString, float.TryParse, onParsed,
                () => type.TryParse<byte, TResult>(valueString, byte.TryParse, onParsed,
                () => type.TryParse<long, TResult>(valueString, long.TryParse, onParsed,
                () => type.TryParse<DateTime, TResult>(valueString, DateTime.TryParse, onParsed,
                onNotParsed)))))))));
        }

        private delegate bool FuncOut<T>(string valueString, out T value);

        private static TResult TryParse<T, TResult>(this Type type, string valueString, FuncOut<T> callback,
            Func<object, TResult> onParsed,
            Func<TResult> onNotParsed)
        {
            if (type.GUID == typeof(T).GUID)
                if (callback(valueString, out T value))
                    return onParsed(value);

            return onNotParsed();
        }

        public static TResult ParseXlsx<TResource, TResult>(this HttpRequestMessage request,
                IProvideUrl urlHelper,
                Stream xlsx,
                Func<TResource, KeyValuePair<string, string>[], Task<HttpResponseMessage>> executePost,
                Func<TResource, KeyValuePair<string, string>[], Task<HttpResponseMessage>> executePut,
            Func<HttpResponseMessage, TResult> onComplete)
            where TResource : IReferenceable
        {
            return OpenXmlWorkbook.Read(xlsx,
                (workbook) =>
                {
                    return workbook.ReadCustomValues(
                        (customValues) =>
                        {
                            var rowsFromAllSheets = workbook.ReadSheets()
                                .SelectMany(
                                    sheet =>
                                    {
                                        var rows = sheet
                                            .ReadRows()
                                            .ToArray();
                                        if (!rows.Any())
                                            return rows;
                                        return rows.Skip(1);
                                    }).ToArray();
                            return request.ParseXlsxBackground(urlHelper, customValues, rowsFromAllSheets,
                                executePost, executePut, onComplete);
                        });
                });
        }

        private static TResult ParseXlsxBackground<TResource, TResult>(this HttpRequestMessage request,
                IProvideUrl urlHelper,
                KeyValuePair<string, string>[] customValues, string[][] rows,
                Func<TResource, KeyValuePair<string, string>[], Task<HttpResponseMessage>> executePost,
                Func<TResource, KeyValuePair<string, string>[], Task<HttpResponseMessage>> executePut,
            Func<HttpResponseMessage, TResult> onComplete)
            where TResource : IReferenceable
        {
            throw new NotImplementedException();
            //var response = request.CreateResponsesBackground(urlHelper,
            //    (updateProgress) =>
            //    {
            //        var propertyOrder = typeof(TResource)
            //            .GetProperties()
            //            .OrderBy(propInfo =>
            //                propInfo.GetCustomAttribute(
            //                    (SheetColumnAttribute sheetColumn) => sheetColumn.GetSortValue(propInfo),
            //                    () => propInfo.Name));

            //        return rows
            //            .Select(
            //                async (row) =>
            //                {
            //                    var resource = propertyOrder
            //                        .Aggregate(Activator.CreateInstance<TResource>(),
            //                            (aggr, property, index) =>
            //                            {
            //                                var value = row.Length > index ?
            //                                    row[index] : default(string);
            //                                TryCastFromXlsSerialization(property, value,
            //                                    (valueCasted) =>
            //                                    {
            //                                        property.SetValue(aggr, valueCasted);
            //                                        return true;
            //                                    },
            //                                    () => false);
            //                                return aggr;
            //                            });
            //                    if (resource.Id.IsEmpty())
            //                    {
            //                        resource.Id = Guid.NewGuid();
            //                        var postResponse = await executePost(resource, customValues);
            //                        return updateProgress(postResponse);
            //                    }
            //                    var putResponse = await executePut(resource, customValues);
            //                    return updateProgress(putResponse);
            //                })
            //           .WhenAllAsync(10);
            //    },
            //    rows.Length);
            //return onComplete(response);
        }

        public static TResult ParseXlsx<TResult>(this HttpRequestMessage request,
                Stream xlsx, Type[] sheetTypes,
                Func<KeyValuePair<string, string>[], KeyValuePair<string, IReferenceable[]>[], TResult> execute)
        {
            var result = OpenXmlWorkbook.Read(xlsx,
                (workbook) =>
                {
                    var propertyOrders = sheetTypes.Select(type => type.PairWithValue(type.GetProperties().Reverse().ToArray())).ToArray();
                    var x = workbook.ReadCustomValues(
                        (customValues) =>
                        {
                            var resourceList = workbook.ReadSheets()
                                .SelectReduce(
                                    (sheet, next, skip) =>
                                    {
                                        var rows = sheet
                                            .ReadRows()
                                            .ToArray();
                                        if (!rows.Any())
                                            return skip();

                                        var resources = rows
                                            .Skip(1)
                                            .Select(
                                                row =>
                                                {
                                                    var propertyOrderOptions = propertyOrders
                                                        .Where(
                                                            kvp =>
                                                            {
                                                                var ids = kvp.Value
                                                                    .Select((po, index) => index.PairWithValue(po))
                                                                    .Where(poKvp => poKvp.Value.Name == "Id");
                                                                if (!ids.Any())
                                                                    return false;
                                                                var id = ids.First();
                                                                if (row.Length <= id.Key)
                                                                    return false;
                                                                if (!Uri.TryCreate(row[id.Key], UriKind.RelativeOrAbsolute, out Uri urn))
                                                                    return false;
                                                                if (!urn.TryParseUrnNamespaceString(out string[] nss, out string nid))
                                                                {
                                                                    var sections = row[id.Key].Split(new char[] { '/' });
                                                                    if (!sections.Any())
                                                                        return false;
                                                                    nid = sections[0];
                                                                }
                                                                return nid == kvp.Key.Name;
                                                            })
                                                        .ToArray();
                                                    var propertyOrder = propertyOrderOptions.Any() ?
                                                        propertyOrderOptions.First()
                                                        :
                                                        typeof(IReferenceable).PairWithValue(typeof(IReferenceable).GetProperties().Reverse().ToArray());

                                                    var resource = (IReferenceable)propertyOrder.Value
                                                        .Aggregate(Activator.CreateInstance(propertyOrder.Key),
                                                            (aggr, property, index) =>
                                                            {
                                                                var value = row.Length > index ?
                                                                    row[index] : default(string);
                                                                TryCastFromXlsSerialization(property, value,
                                                                    (valueCasted) =>
                                                                    {
                                                                        property.SetValue(aggr, valueCasted);
                                                                        return true;
                                                                    },
                                                                    () => false);
                                                                return aggr;
                                                            });
                                                    return resource;
                                                })
                                            .ToArray();
                                        return next(sheet.Name.PairWithValue(resources));
                                    },
                                    (KeyValuePair<string, IReferenceable[]>[] resourceLists) =>
                                    {
                                        return execute(customValues, resourceLists);
                                    });
                            return resourceList;
                        });
                    return x;
                });
            return result;
        }

        public static async Task<TResult> ParseXlsxAsync<TResource, TResult>(this HttpRequestMessage request,
                Stream xlsx,
                Func<TResource, KeyValuePair<string, string>[], Task<HttpResponseMessage>> executePost,
                Func<TResource, KeyValuePair<string, string>[], Task<HttpResponseMessage>> executePut,
            Func<HttpResponseMessage[], TResult> onComplete)
            where TResource : IReferenceable
        {
            var result = await OpenXmlWorkbook.Read(xlsx,

            async (workbook) =>
            {
                return await workbook.ReadCustomValues(
                    async (customValues) =>
                    {
                        var propertyOrder = typeof(TResource).GetProperties();
                        var x = await workbook.ReadSheets()
                            .Select(
                                sheet =>
                                {
                                    var rows = sheet
                                        .ReadRows()
                                        .ToArray();
                                    if (!rows.Any())
                                        return EnumerableAsync.Empty<HttpResponseMessage>();

                                    return rows
                                        .Skip(1)
                                        .Select(
                                            row =>
                                            {
                                                var resource = propertyOrder
                                                    .Aggregate(Activator.CreateInstance<TResource>(),
                                                        (aggr, property, index) =>
                                                        {
                                                            var value = row.Length > index ?
                                                                row[index] : default(string);
                                                            TryCastFromXlsSerialization(property, value,
                                                                (valueCasted) =>
                                                                {
                                                                    property.SetValue(aggr, valueCasted);
                                                                    return true;
                                                                },
                                                                () => false);
                                                            return aggr;
                                                        });
                                                return executePut(resource, customValues);
                                            })
                                        .AsyncEnumerable(10);
                                })
                           .SelectMany()
                           .ToArrayAsync();
                        return onComplete(x);
                    });
            });
            return result;
        }

        #endregion

    }
}
