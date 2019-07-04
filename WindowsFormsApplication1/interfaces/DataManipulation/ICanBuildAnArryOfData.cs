using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTranslator.interfaces.DataManipulation
{
    interface ICanBuildAnArryOfData
    {
        object[] buildArray(List<T> data);
    }
}
