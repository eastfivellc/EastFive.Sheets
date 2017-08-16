using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EastFive.Sheets
{
    class OpenXmlSheet : ISheet
    {
        private Worksheet worksheetData;

        public OpenXmlSheet(Worksheet worksheetData)
        {
            this.worksheetData = worksheetData;
        }

        public IEnumerable<string[]> ReadRows()
        {
            var rows = worksheetData
                .Descendants<Row>()
                .Select(row => row
                    .Elements<Cell>()
                    .Select(
                        (cell) => cell.CellValue.Text)
                    .ToArray())
                ;
            return rows;
        }
    }
}
