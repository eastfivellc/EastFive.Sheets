using System;
using ClosedXML.Excel;
using System.Reflection;

namespace EastFive.Sheets
{
    public interface IWriteCell
    {
        IHeaderData ComputeHeader(Type resourceType, int colIndex, MemberInfo member, Type[] otherTypes);
        IXLCell WriteHeader<TResource>(IXLWorksheet worksheet, IHeaderData headerDataObj, Type[] otherTypes, TResource[] resources);
        IXLCell WriteCell<TResource>(IXLWorksheet worksheet, IHeaderData headerData, int rowIndex, TResource resource);
    }
}

