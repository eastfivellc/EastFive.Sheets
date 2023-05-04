using System;
using ClosedXML.Excel;
using EastFive.Sheets.Api;

namespace EastFive.Sheets
{
    public interface IWriteSheet
    {
        string GetSheetName(Type resourceType);
        IHeaderData[] GetHeaderData(Type resourceType, Type[] otherTypes);
        void WriteSheet<TResource>(IXLWorksheet worksheet,
            TResource[] resource1s, Type[] types, IHeaderData[] headerDatas);
    }

}

