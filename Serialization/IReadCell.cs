using System;
using ClosedXML.Excel;
using System.Reflection;

namespace EastFive.Sheets
{
    public interface IMatchHeaderData
    {

    }

    public interface IReadCell
    {
        IMatchHeaderData MatchHeaders<TResource>(IXLRow headerRow, MemberInfo memberInfo);
        TResource ReadCell<TResource>(TResource resource, IXLRow headerRow,
            IMatchHeaderData headerData);
    }
}

