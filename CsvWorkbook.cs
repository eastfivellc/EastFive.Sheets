using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CsvHelper;
using EastFive.Extensions;

namespace EastFive.Sheets
{
    public class CsvWorkbook : IUnderstandSheets
    {
        private readonly CsvSheet sheet;

        public CsvWorkbook(Stream stream)
        {
            this.sheet = new CsvSheet(stream);
        }

        public void Dispose()
        {
            
        }

        public TResult ReadCustomValues<TResult>(
            Func<KeyValuePair<string, string>[], TResult> onResults)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<ISheet> ReadSheets()
        {
            return sheet.AsEnumerable();
        }

        public TResult WriteCustomProperties<TResult>(Func<Action<string, string>, TResult> callback)
        {
            throw new NotImplementedException();
        }

        public TResult WriteSheetByRow<TResult>(
            Func<Action<object[]>, TResult> p,
            string sheetName = "Sheet1")
        {
            throw new NotImplementedException();
        }
    }
}
