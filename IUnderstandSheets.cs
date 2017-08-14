using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EastFive.Sheets
{
    public interface IUnderstandSheets
    {
        TResult ReadCustomValues<TResult>(Func<KeyValuePair<string, string>[], TResult> onResults);
        IEnumerable<ISheet> ReadSheets();
    }
}
