using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CsvHelper;

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
            List<string[]> rows = new List<string[]>();
            using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(stream))
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
