using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mime;
using System.Threading.Tasks;

using Azure.Storage.Blobs.Models;

using CsvHelper;

using EastFive.Persistence;
using EastFive.Persistence.Azure;
using EastFive.Azure.Persistence.StorageTables;
using EastFive.Azure.Persistence.AzureStorageTables;
using EastFive.Azure.StorageTables;
using EastFive.Azure.Persistence;
using EastFive.Azure.Persistence.Blobs;
using EastFive.Serialization;

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
                        await blobStream.WriteCSVAsync(csvData, leaveOpen: true);
                    }
                    catch (Exception ex)
                    {
                        var msgBytes = ex.Message.GetBytes(System.Text.Encoding.UTF8);
                        await blobStream.WriteAsync(msgBytes, 0, msgBytes.Length, default);
                    }
                },
                onSuccess:(blobInfo) => blobInfo,
                contentDisposition: contentDisposition,
                contentType: contentType,
                connectionStringConfigKey: EastFive.Azure.AppSettings.Persistence.DataLake.ConnectionString);
        }

        public static async Task<BlobContentInfo> StorageSaveCSV(this IEnumerable<string[]> csvData,
            AzureBlobFileSystemUri filePathAndName, ContentDisposition contentDisposition, ContentType contentType)
        {
            return await filePathAndName.BlobCreateOrUpdateAsync(
                writeStreamAsync: async (blobStream) =>
                {
                    try
                    {
                        await blobStream.WriteCSVAsync(csvData, leaveOpen: true);
                    }
                    catch (Exception ex)
                    {
                        var msgBytes = ex.Message.GetBytes(System.Text.Encoding.UTF8);
                        await blobStream.WriteAsync(msgBytes, 0, msgBytes.Length, default);
                    }
                },
                onSuccess: (blobInfo) => blobInfo,
                contentDisposition: contentDisposition,
                contentType: contentType);
        }
    }
}

