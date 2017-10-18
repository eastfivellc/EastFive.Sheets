using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace EastFive.Sheets
{
    class AdoXlsSheetERF : ISheet
    {
        private DataTable dataTable;

        public AdoXlsSheetERF(DataTable dataTable)
        {
            this.dataTable = dataTable;
        }

        public IEnumerable<string[]> ReadRows()
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
    }
}
