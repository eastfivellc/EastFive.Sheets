using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EastFive.Sheets
{
    public interface ISheet
    {
        IEnumerable<string[]> ReadRows();
    }
}
