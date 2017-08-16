using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EastFive.Sheets
{
    public interface IUnderstandSheets : IDisposable
    {
        TResult ReadCustomValues<TResult>(Func<KeyValuePair<string, string>[], TResult> onResults);
        IEnumerable<ISheet> ReadSheets();
        TResult WriteCustomProperties<TResult>(
            Func<Action<string, string>, TResult> callback);
        TResult WriteSheetByRow<TResult>(Func<Action<object[]>, TResult> p);
    }
}
