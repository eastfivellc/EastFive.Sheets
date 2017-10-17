using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EastFive.Sheets
{
    public class CsvSheet : ISheet
    {
        private readonly Stream fileStream;

        public CsvSheet(Stream fileStream)
        {
            this.fileStream = fileStream;    
        }

        public IEnumerable<string[]> ReadRows()
        {
            fileStream.Seek(0, SeekOrigin.Begin);
            List<string[]> rows = new List<string[]>();
            var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(fileStream);
            parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
            parser.SetDelimiters(",");
            while (!parser.EndOfData)
            {
                string[] fields = new string[0];
                try
                {
                    fields = parser.ReadFields(); 
                }
                catch (Exception ex)
                {
                    continue;
                }
                rows.Add(fields);
            }
            parser.Close();
            return rows.ToArray();
        }
    }
}
