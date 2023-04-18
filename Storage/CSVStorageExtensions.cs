using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using CsvHelper;

using EastFive.Persistence;
using EastFive.Persistence.Azure;
using EastFive.Azure.Persistence.StorageTables;
using EastFive.Azure.Persistence.AzureStorageTables;
using EastFive.Azure.StorageTables;
using EastFive.Azure.Persistence;
using EastFive.Azure.Persistence.Blobs;
using Azure.Storage.Blobs.Models;
using System.Net.Mime;

namespace EastFive.Sheets.Storage
{
	public static class CSVStorageExtensions
	{
		public static async Task<BlobContentInfo> StorageSaveCSV(this IEnumerable<string[]> csvData,
            string name, string containerName,
            ContentDisposition contentDisposition, ContentType contentType)
		{
            return await name.BlobCreateOrUpdateAsync(containerName,
                writeStreamAsync:async (blobStream) =>
                {
                    try
                    {
                        using (TextWriter writer = new StreamWriter(blobStream, System.Text.Encoding.UTF8, leaveOpen:true))
                        {
                            using (var csv = new CsvWriter(writer, System.Globalization.CultureInfo.InvariantCulture, leaveOpen:true))
                            {
                                foreach (var row in csvData)
                                {
                                    foreach (var col in row)
                                        csv.WriteConvertedField(col); // where values implements IEnumerable
                                    await csv.NextRecordAsync();
                                    await writer.FlushAsync();
                                    await blobStream.FlushAsync();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                },
                onSuccess:(blobInfo) => blobInfo,
                contentDisposition: contentDisposition,
                contentType: contentType,
                connectionStringConfigKey: EastFive.Azure.AppSettings.Persistence.DataLake.ConnectionString);
        }
	}
}

