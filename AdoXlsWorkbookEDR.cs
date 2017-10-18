using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace EastFive.Sheets
{
    public class AdoXlsWorkbookEDR : IUnderstandSheets
    {
        private DataSet data;

        public AdoXlsWorkbookEDR()
        {

        }

        public AdoXlsWorkbookEDR(DataSet data)
        {
            this.data = data;
        }

        public static AdoXlsWorkbookEDR Load(byte [] xlsFile)
        {
            using (MemoryStream stream = new MemoryStream(xlsFile))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var data = reader.AsDataSet();
                    return new AdoXlsWorkbookEDR(data);
                }
            }
        }

        public void Dispose()
        {
            data.Dispose();
        }

        public TResult ReadCustomValues<TResult>(Func<KeyValuePair<string, string>[], TResult> onResults)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<ISheet> ReadSheets()
        {
            if (data.Tables == null || data.Tables.Count == 0)
                yield break;

            foreach(DataTable table in data.Tables)
            {
                yield return new AdoXlsSheetEDR(table);
            }
        }

        public TResult WriteCustomProperties<TResult>(Func<Action<string, string>, TResult> callback)
        {
            throw new NotImplementedException();
        }

        public TResult WriteSheetByRow<TResult>(Func<Action<object[]>, TResult> p)
        {
            throw new NotImplementedException();
        }
    }
}
