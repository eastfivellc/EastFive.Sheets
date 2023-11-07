using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CsvHelper;
using EastFive.Linq;
using static EastFive.Azure.Persistence.AzureStorageTables.Backups.TableBackupOperation;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace EastFive.Sheets
{
    public class CsvSheet : ISheet
    {
        private readonly Stream stream;

        public CsvSheet(Stream stream)
        {
            this.stream = stream;    
        }

        public string Name => "sheet";

        public IEnumerable<string[]> ReadRows(Func<Type, object, Func<string>, string> serializer = default)
        {
            stream.Seek(0, SeekOrigin.Begin);
            var rawData = stream.ToBytes();
            var encoding = DecodeEncoding(rawData);
            List<string[]> rows = new List<string[]>();
            using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(
                new MemoryStream(rawData), encoding))
            {
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                parser.SetDelimiters(",");
                while (!parser.EndOfData)
                {
                    string[] fields = new string[0];
                    try
                    {
                        fields = parser.ReadFields();
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    rows.Add(fields);
                }
            }
            return rows.ToArray();
        }

        public Encoding DecodeEncoding(byte[] rawData)
        {
            var encodingProfiles = Encoding
                .GetEncodings()
                .Select(encodingInfo => Encoding.GetEncoding(encodingInfo.CodePage))
                .Select(
                    encoding =>
                    {
                        var rowSizes = new List<double>();
                        bool didThrowException = false;
                        using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(
                            new MemoryStream(rawData), encoding))
                        {
                            parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                            parser.SetDelimiters(",");
                            while (!parser.EndOfData)
                            {
                                string[] fields = new string[0];
                                try
                                {
                                    fields = parser.ReadFields();
                                }
                                catch (Exception)
                                {
                                    didThrowException = true;
                                    continue;
                                }
                                rowSizes.Add(fields.Length);
                            }
                        }

                        var rowCount = rowSizes.Count;
                        if (rowCount == 0)
                            return (
                                encoding,
                                rowCount,
                                rowSizeAvg: 0d,
                                stdDev: 0d,
                                stdError: 0d,
                                totalCells: 0,
                                true);

                        var rowSizeAvg = rowSizes.Average();
                        var stdDev = rowSizes.StdDev();
                        var totalCells = rowSizes.Sum();
                        return (
                            encoding,
                            rowCount,
                            rowSizeAvg: rowSizeAvg,
                            stdDev: stdDev,
                            stdError: stdDev / Math.Sqrt(rowCount),
                            totalCells: totalCells,
                            didThrowException);
                    })
                .Where(ep => ep.rowCount > 0)
                .Where(ep => ep.rowSizeAvg > 0)
                .Where(ep => !ep.didThrowException)
                .ToArray();

            if (encodingProfiles.None())
                return Encoding.Default;

            var averageCellCount = encodingProfiles
                .Average(tpl => Math.Sqrt(tpl.totalCells));

            var meanStdError = encodingProfiles
                .Select(tpl => tpl.stdError)
                .Average();

            var meanRowSizeAvg = encodingProfiles
                .Select(tpl => tpl.rowSizeAvg)
                .Average();

            var meanRows = encodingProfiles
                .Select(tpl => tpl.rowCount)
                .Average();

            var bySquareness = encodingProfiles
                .Select(
                    ep =>
                    {
                        var z = Math.Log(ep.rowCount);
                        var y = Math.Log(ep.rowSizeAvg);
                        var d = Math.Sqrt((z * z) + (y * y));
                        var v = d - ep.stdError;
                        return (
                            rank: v,
                            squareness: d,
                            ep.stdError,
                            averageColumns: ep.rowSizeAvg,
                            totalRows: ep.rowCount,
                            ep.encoding);
                    })
                .OrderByDescending(ep => ep.rank)
                .ToArray();

            return bySquareness.First().encoding;

            //Encoding SelectBest((Encoding encoding, int count, double mean, double stdDev, double stdError, double totalCells, double squareness, bool didThrowException)[] encodingProfiles)
            //{

            //}



        }

        public void WriteRows(string fileName, object[] rows)
        {
            // Note that the CSVHelper library expects the properties in the incoming object[] to be in the 
            // format of 
            //public class Foo
            //{
            //    public string Id { get; set; }
            //    public string Thing1 { get; set; }
            //}
            //The properties must be public and must have a getter and setter

            using (var textWriter = File.CreateText(fileName))
            using (var writer = new CsvWriter(textWriter, System.Globalization.CultureInfo.InvariantCulture))
            {
                writer.WriteRecords(rows);
            }
        }

        public void WriteRows<T>(IEnumerable<T> rows, bool leaveOpen = false)
        {
            // Note that the CSVHelper library expects the properties in the incoming object[] to be in the 
            // format of 
            //public class Foo
            //{
            //    public string Id { get; set; }
            //    public string Thing1 { get; set; }S
            //}
            //The properties must be public and must have a getter and setter
            using (var streamWriter = new StreamWriter(stream, new UTF8Encoding(false), 4096, leaveOpen))
            using (var writer = new CsvWriter(streamWriter, System.Globalization.CultureInfo.InvariantCulture))
            {
                writer.WriteRecords(rows);
            }
        }
    }
}
