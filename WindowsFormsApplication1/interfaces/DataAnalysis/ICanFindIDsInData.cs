using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTranslator.interfaces.DataAnalysis
{
    interface ICanFindIDsInData
    {
        string FindID(object[] data);
    }
}
