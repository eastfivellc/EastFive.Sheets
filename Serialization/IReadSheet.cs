using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using EastFive.Sheets.Api;

namespace EastFive.Sheets
{
    public interface IReadSheet
    {
        string GetSheetName(Type resourceType);
        IEnumerable<TResource> ReadSheet<TResource>(IXLWorksheet worksheet,
            IEnumerable<IXLRow> rows);
    }

}

