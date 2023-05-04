using System;
using EastFive.Sheets.Api;
using System.Reflection;

namespace EastFive.Sheets
{
    public interface IHeaderData
    {
        int ColumnsIndex { get; }

        int ColumnsUsed { get; }

        IWriteCell CellWriter { get; }

        MemberInfo MemberInfo { get; }
    }
}

