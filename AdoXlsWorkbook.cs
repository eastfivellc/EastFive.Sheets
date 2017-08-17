using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EastFive.Sheets
{
    public class AdoXlsWorkbook : IUnderstandSheets
    {
        private OleDbConnection conn;
        private string filename;

        public AdoXlsWorkbook()
        {

        }

        public AdoXlsWorkbook(OleDbConnection conn, string filename)
        {
            this.conn = conn;
            this.filename = filename;
            conn.Open();
        }

        public static AdoXlsWorkbook Load(byte [] xlsFile)
        {
            var filename = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(filename, xlsFile);
            //var connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties=\"Excel 12.0;HDR=YES\"";
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"", filename);

            // Do your work on excel
            var conn = new System.Data.OleDb.OleDbConnection(connectionString);
            return new AdoXlsWorkbook(conn, filename);
        }

        public void Dispose()
        {
            conn.Close();
            conn.Dispose();
            System.IO.File.Delete(filename);
        }

        public TResult ReadCustomValues<TResult>(Func<KeyValuePair<string, string>[], TResult> onResults)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<ISheet> ReadSheets()
        {
            var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt == null)
                yield break;
            
            // Add the sheet name to the string array.
            foreach(DataRow row in dt.Rows)
            {
                var sheetName = row["TABLE_NAME"].ToString();
                yield return new AdoXlsSheet(conn, sheetName);
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
