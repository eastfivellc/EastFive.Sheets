using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EastFive.Sheets
{
    class AdoXlsSheetEDR : ISheet
    {
        private DataTable dataTable;

        public AdoXlsSheetEDR(DataTable dataTable)
        {
            this.dataTable = dataTable;
        }

        public string Name => dataTable.TableName;

        public IEnumerable<string[]> ReadRows(Func<Type, object,
            Func<string>, string> discardSerializer = default,
            bool discardAutoDecodeEncoding = default,
            Encoding discardEncodingToUse = default)
        {
            foreach(DataRow row in dataTable.Rows)
            {
                yield return Enumerable
                    .Range(0, dataTable.Columns.Count)
                    .Select(
                        columnIndex =>
                        {
                            var x = row[columnIndex];
                            if (typeof(string).IsInstanceOfType(x))
                                return x.ToString();
                            if (typeof(DBNull).IsInstanceOfType(x))
                                return string.Empty;
                            return x.ToString();
                        })
                    .ToArray();
            }
        }

        public void WriteRows(string fileName, object[] rows)
        {
            throw new NotImplementedException();
        }
    }
}
