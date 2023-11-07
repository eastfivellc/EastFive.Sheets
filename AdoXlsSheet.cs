using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EastFive.Sheets
{
    class AdoXlsSheet : ISheet
    {
        private OleDbConnection conn;
        private string sheetName;

        public AdoXlsSheet(OleDbConnection conn, string sheetName)
        {
            this.conn = conn;
            this.sheetName = sheetName;
        }

        public string Name => sheetName;

        public IEnumerable<string[]> ReadRows(Func<Type, object, Func<string>, string> serializer = default)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT * FROM [{sheetName}]";
                using (var rdr = cmd.ExecuteReader())
                {
                    int rowNumber = 0;
                    while (rdr.Read())
                    {
                        yield return Enumerable
                            .Range(0, rdr.FieldCount)
                            .Select(
                                columnIndex =>
                                {
                                    var x = rdr[columnIndex];
                                    if(typeof(string).IsInstanceOfType(x))
                                        return x.ToString();
                                    if (typeof(DBNull).IsInstanceOfType(x))
                                        return string.Empty;
                                    return x.ToString();
                                })
                            .ToArray();
                        rowNumber++;
                    }
                }
            }
        }

        public void WriteRows(string fileName, object[] rows)
        {
            throw new NotImplementedException();
        }
    }
}
