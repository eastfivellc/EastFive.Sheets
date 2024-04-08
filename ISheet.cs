using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EastFive.Sheets
{
    public interface ISheet
    {
        IEnumerable<string[]> ReadRows(Func<Type, object,
            Func<string>, string> serializer = default,
            bool autoDecodeEncoding = default,
            Encoding encodingToUse = default);

        void WriteRows(string fileName, object[] rows);

        string Name { get; }
    }
}
